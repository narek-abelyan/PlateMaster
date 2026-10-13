"""Portable analysis settings and provenance, independent of the Streamlit UI."""
from __future__ import annotations

from copy import deepcopy
import hashlib
from importlib import metadata
import json
import math
from pathlib import Path
import platform
import subprocess
from typing import Any

import pandas as pd

APP_VERSION = "1.3.0-dev"
SCHEMA_VERSION = 1
MAX_SETTINGS_BYTES = 2 * 1024 * 1024
AGGREGATION_NAMES = ["Mean", "Median", "Min", "Max", "Std", "Sum"]
LAYOUTS = ["By original columns", "By original rows", "By plate groups (template)"]
PLOT_TYPES = [
    "Line + points", "Line + points (по Columns)", "Scatter", "Bar plot (по Time)",
    "Bar plot (по Columns)", "Box plot (по Time)", "Box plot (по Columns)",
    "Violin plot (по Time)", "Violin plot (по Columns)",
]
DEFAULT_SETTINGS = {
    "agg_layout_mode": LAYOUTS[0], "selected_rows": [], "selected_columns": [],
    "agg_name": "Mean", "outlier_method": "None", "iqr_multiplier": 1.5,
    "zscore_threshold": 3.0, "mad_threshold": 3.5, "color_scale_mode": "Global scale",
    "color_palette": "Rose", "plot_columns": [], "plot_type": PLOT_TYPES[0],
    "error_method": "None", "stat_time": None, "stat_metric": None,
    "ttest_time_a": None, "ttest_metric_a": None, "ttest_time_b": None,
    "ttest_metric_b": None,
}
ENUMS = {
    "agg_layout_mode": LAYOUTS, "agg_name": AGGREGATION_NAMES,
    "outlier_method": ["None", "IQR", "Z-score", "MAD"],
    "color_scale_mode": ["Global scale", "Per-column scale", "Per row scale"],
    "color_palette": ["Rose", "Pink-Violet", "Purple", "Blue", "Green", "Orange", "Teal", "Gray"],
    "plot_type": PLOT_TYPES, "error_method": ["None", "Std", "SEM", "95% CI", "IQR/2", "MAD"],
}
THRESHOLDS = {"iqr_multiplier": (0.5, 5.0), "zscore_threshold": (1.0, 5.0), "mad_threshold": (1.0, 8.0)}


def dumps_config(config: dict) -> bytes:
    """Encode portable JSON without non-standard NaN/Infinity values."""
    return json.dumps(config, ensure_ascii=False, allow_nan=False, indent=2, sort_keys=True).encode("utf-8")


def validate_config(config: Any) -> dict:
    """Validate a settings document before any widget state can be changed."""
    if not isinstance(config, dict) or config.get("format") != "platemaster.analysis":
        raise ValueError("Not a PlateMaster analysis settings file.")
    if type(config.get("schema_version")) is not int or config["schema_version"] != SCHEMA_VERSION:
        raise ValueError("Unsupported settings schema version.")
    settings = config.get("settings")
    if not isinstance(settings, dict) or set(settings) != set(DEFAULT_SETTINGS):
        raise ValueError("Settings fields are missing or unknown.")
    for key, choices in ENUMS.items():
        if settings[key] not in choices:
            raise ValueError(f"Unsupported {key}: {settings[key]!r}.")
    for key in ("selected_rows", "selected_columns", "plot_columns"):
        value = settings[key]
        if not isinstance(value, list) or not all(isinstance(x, str) for x in value):
            raise ValueError(f"{key} must be a list of names.")
        if len(value) != len(set(value)):
            raise ValueError(f"{key} contains duplicate names.")
    for key, (low, high) in THRESHOLDS.items():
        value = settings[key]
        if type(value) not in (int, float) or not math.isfinite(value) or not low <= value <= high:
            raise ValueError(f"{key} must be between {low} and {high}.")
    for key in ("stat_time", "ttest_time_a", "ttest_time_b"):
        value = settings[key]
        if value is not None and (type(value) not in (int, float) or not math.isfinite(value)):
            raise ValueError(f"{key} must be a finite number or null.")
    for key in ("stat_metric", "ttest_metric_a", "ttest_metric_b"):
        if settings[key] is not None and not isinstance(settings[key], str):
            raise ValueError(f"{key} must be a name or null.")
    if settings["agg_name"] != "Mean" and settings["error_method"] in ("SEM", "95% CI"):
        raise ValueError("SEM and 95% CI error bars require Mean aggregation.")
    groups = config.get("plate_groups")
    if not isinstance(groups, list):
        raise ValueError("plate_groups must be a list of labelled cells.")
    seen = set()
    for cell in groups:
        if not isinstance(cell, dict) or set(cell) != {"row_id", "column", "group"}:
            raise ValueError("Each plate-group cell needs row_id, column and group.")
        if not all(isinstance(v, str) and v.strip() for v in cell.values()):
            raise ValueError("Plate-group labels must be non-empty text.")
        address = (cell["row_id"], cell["column"])
        if address in seen:
            raise ValueError("A plate cell cannot belong to more than one group.")
        seen.add(address)
        if cell["row_id"] not in settings["selected_rows"] or cell["column"] not in settings["selected_columns"]:
            raise ValueError("Plate-group cells must belong to the selected rows and columns.")
    metrics = metric_names(settings, groups)
    for key in ("stat_metric", "ttest_metric_a", "ttest_metric_b"):
        if settings[key] is not None and settings[key] not in metrics:
            raise ValueError(f"{key} is not present in this analysis layout.")
    if not set(settings["plot_columns"]).issubset(metrics):
        raise ValueError("Plot columns are not present in this analysis layout.")
    inputs = config.get("inputs")
    if not isinstance(inputs, list):
        raise ValueError("inputs must be a list.")
    names = set()
    for item in inputs:
        if not isinstance(item, dict) or not isinstance(item.get("file"), str):
            raise ValueError("Each input record needs a file name.")
        if item["file"] in names:
            raise ValueError("Input file names must be unique.")
        names.add(item["file"])
        for key in ("sha256", "table_sha256"):
            value = item.get(key)
            if value is not None and (not isinstance(value, str) or len(value) != 64 or any(c not in "0123456789abcdef" for c in value)):
                raise ValueError(f"Invalid input {key}.")
    if not isinstance(config.get("software"), dict):
        raise ValueError("software must be an object.")
    # Also rejects non-finite numbers hidden in provenance rather than widget fields.
    dumps_config(config)
    return deepcopy(config)


def loads_config(content: bytes) -> dict:
    """Read a bounded JSON file; no code, paths, or objects are executed."""
    if len(content) > MAX_SETTINGS_BYTES:
        raise ValueError("Settings file is too large (maximum 2 MiB).")
    try:
        config = json.loads(content.decode("utf-8-sig"))
        return validate_config(config)
    except (UnicodeDecodeError, json.JSONDecodeError, RecursionError) as exc:
        raise ValueError("Invalid settings JSON.") from exc


def metric_names(settings: dict, groups: list[dict]) -> list[str]:
    """Return the measurement labels produced by the selected layout."""
    if settings["agg_layout_mode"] == LAYOUTS[1]:
        return list(settings["selected_rows"])
    if settings["agg_layout_mode"] == LAYOUTS[2]:
        return sorted({c["group"] for c in groups})
    return list(settings["selected_columns"])


def make_config(settings: dict, groups: list[dict], inputs: list[dict], software: dict) -> dict:
    """Build a complete, validated analysis snapshot."""
    return validate_config({"format": "platemaster.analysis", "schema_version": SCHEMA_VERSION,
                            "settings": settings, "plate_groups": groups, "inputs": inputs,
                            "software": software})


def table_fingerprint(table: pd.DataFrame) -> str:
    """Hash the parsed table as a second check on spreadsheet-reader behaviour."""
    return hashlib.sha256(table.to_json(orient="split", double_precision=15, index=False).encode("utf-8")).hexdigest()


def input_record(name: str, content: bytes | None, table: pd.DataFrame, time: float | None) -> dict:
    """Record exact input bytes and the table actually used for analysis."""
    return {"file": name, "sha256": hashlib.sha256(content).hexdigest() if content is not None else None,
            "size_bytes": len(content) if content is not None else None,
            "table_sha256": table_fingerprint(table), "time": time,
            "time_source": "last numeric token in filename", "time_unit": "unspecified"}


def software_record(directory: Path) -> dict:
    """Identify both the Git revision and the actual source files, including uncommitted edits."""
    libraries = {}
    for name in ("streamlit", "pandas", "numpy", "scipy", "altair", "openpyxl", "xlrd", "python-calamine"):
        try:
            libraries[name] = metadata.version(name)
        except metadata.PackageNotFoundError:
            libraries[name] = "not installed"
    result = {"application_version": APP_VERSION, "python": platform.python_version(),
              "platform": platform.platform(), "libraries": libraries,
              "source_sha256": {p.name: hashlib.sha256(p.read_bytes()).hexdigest()
                                for p in sorted(directory.glob("*.py"))},
              "git_commit": None, "working_tree_modified": None}
    try:
        commit = subprocess.run(["git", "-C", str(directory), "rev-parse", "HEAD"], capture_output=True, text=True, timeout=3)
        if commit.returncode == 0:
            result["git_commit"] = commit.stdout.strip()
            status = subprocess.run(["git", "-C", str(directory), "status", "--porcelain"], capture_output=True, text=True, timeout=3)
            if status.returncode == 0:
                result["working_tree_modified"] = bool(status.stdout.strip())
    except (OSError, subprocess.TimeoutExpired):
        pass
    return result


def replay_warnings(saved: dict, current_inputs: list[dict], current_software: dict) -> list[str]:
    """Describe differences without treating reuse on a new dataset as exact reproduction."""
    messages = []
    expected = {x["file"]: x for x in saved["inputs"]}
    current = {x["file"]: x for x in current_inputs}
    for name in sorted(expected.keys() - current.keys()):
        messages.append(f"Missing input: {name}")
    for name in sorted(current.keys() - expected.keys()):
        messages.append(f"Additional input: {name}")
    for name in sorted(expected.keys() & current.keys()):
        old, new = expected[name], current[name]
        if old.get("sha256") is None or new.get("sha256") is None:
            messages.append(f"Original file bytes cannot be verified: {name}")
        elif old["sha256"] != new["sha256"]:
            messages.append(f"Input file changed: {name}")
        if old.get("table_sha256") != new.get("table_sha256"):
            messages.append(f"Parsed measurements changed: {name}")
        if old.get("time") != new.get("time"):
            messages.append(f"Filename time changed: {name}")
    for key in ("application_version", "python", "platform", "libraries", "source_sha256"):
        if saved["software"].get(key) != current_software.get(key):
            messages.append(f"Software/environment differs: {key}")
    return messages


def provenance_sheets(config: dict) -> dict[str, pd.DataFrame]:
    """Human-readable provenance plus the exact JSON, split below Excel's cell limit."""
    payload = dumps_config(config).decode("utf-8")
    environment = []
    for key, value in config["software"].items():
        if isinstance(value, dict):
            environment.extend({"Field": f"{key}.{k}", "Value": v} for k, v in value.items())
        else:
            environment.append({"Field": key, "Value": value})
    return {
        "Analysis settings": pd.DataFrame([{"Setting": key, "Value": json.dumps(value, ensure_ascii=False)}
                                           for key, value in config["settings"].items()]),
        "Plate groups": pd.DataFrame(config["plate_groups"], columns=["row_id", "column", "group"]),
        "Input files": pd.DataFrame(config["inputs"]),
        "Software versions": pd.DataFrame(environment),
        "Settings JSON": pd.DataFrame([{"Part": i // 30000 + 1, "JSON": payload[i:i + 30000]}
                                        for i in range(0, len(payload), 30000)]),
    }
