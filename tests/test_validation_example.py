"""Regression check for the documented CC0 validation dataset in example_2."""
import hashlib
import json
from pathlib import Path

import pandas as pd
import pytest

from analysis_settings import input_record, loads_config


ROOT = Path(__file__).resolve().parents[1]
EXAMPLE = ROOT / "example_2"


def test_documented_validation_example_matches_reference_results(app_module):
    data_files = sorted((EXAMPLE / "data").glob("validation_plate_*.xlsx"))
    assert [path.stem for path in data_files] == [
        "validation_plate_0", "validation_plate_10", "validation_plate_15",
        "validation_plate_20", "validation_plate_5",
    ]
    assert len(data_files) == 5

    config = loads_config((EXAMPLE / "analysis_settings.json").read_bytes())
    manifest = json.loads((EXAMPLE / "manifest.json").read_text(encoding="utf-8"))
    expected = pd.read_csv(EXAMPLE / "expected_group_results.csv")
    expected_welch = pd.read_csv(EXAMPLE / "expected_welch_comparison.csv").iloc[0]
    mapping = pd.read_csv(EXAMPLE / "plate_map.csv", dtype=str)
    group_map = {
        group: list(zip(part["row_id"], part["column"]))
        for group, part in mapping.groupby("group", sort=False)
    }
    recorded_inputs = {item["file"]: item for item in config["inputs"]}
    manifest_inputs = {item["file"]: item for item in manifest["files"]}
    assert manifest["license"] == "CC0-1.0"
    assert manifest["time_unit"] == "minutes"
    assert "CC0-1.0" in (EXAMPLE / "LICENSE.txt").read_text(encoding="utf-8")
    assert config["plate_groups"] == mapping.to_dict(orient="records")
    assert config["settings"]["agg_layout_mode"] == "By plate groups (template)"
    assert config["settings"]["outlier_method"] == "None"
    collected = {}

    for path in data_files:
        table = app_module.load_table_from_path(path)
        time = app_module._extract_time_from_filename(path.name)
        actual_record = input_record(path.name, path.read_bytes(), table, time)
        assert actual_record["sha256"] == recorded_inputs[path.name]["sha256"]
        assert actual_record["sha256"] == manifest_inputs[path.name]["sha256"]
        assert actual_record["sha256"] == hashlib.sha256(path.read_bytes()).hexdigest()
        assert actual_record["table_sha256"] == recorded_inputs[path.name]["table_sha256"]

        aggregated, removed, dispersion, source_values = app_module.aggregate_by_groups(
            table, group_map, app_module.AGGREGATIONS["Mean"], "None"
        )
        assert removed == []
        collected[time] = source_values

        for group in ("Control", "Treated"):
            reference = expected[(expected["time_minutes"] == time) & (expected["group"] == group)].iloc[0]
            assert aggregated.iloc[0][group] == pytest.approx(reference["mean"], rel=1e-12)
            assert dispersion[group]["std"] == pytest.approx(reference["sample_sd"], rel=1e-12)
            assert dispersion[group]["sem"] == pytest.approx(reference["sem"], rel=1e-12)
            assert dispersion[group]["ci95"] == pytest.approx(reference["ci95_half_width"], rel=1e-12)

            metrics = app_module._basic_stat_metrics(source_values[group]).set_index("Metric")["Value"]
            assert metrics["n"] == reference["n"]
            assert metrics["median"] == pytest.approx(reference["median"], rel=1e-12)
            assert metrics["min"] == pytest.approx(reference["minimum"], rel=1e-12)
            assert metrics["max"] == pytest.approx(reference["maximum"], rel=1e-12)
            assert metrics["95% CI lower (mean)"] == pytest.approx(reference["ci95_lower"], rel=1e-12)
            assert metrics["95% CI upper (mean)"] == pytest.approx(reference["ci95_upper"], rel=1e-12)
            assert metrics["p-value (mean != 0, Student t)"] == pytest.approx(
                reference["one_sample_p_two_sided"], rel=1e-12
            )

    welch = app_module._welch_result(collected[0.0]["Control"], collected[20.0]["Treated"])
    assert welch["n_a"] == expected_welch["n_a"]
    assert welch["n_b"] == expected_welch["n_b"]
    assert welch["mean_a"] == pytest.approx(expected_welch["mean_a"], rel=1e-12)
    assert welch["mean_b"] == pytest.approx(expected_welch["mean_b"], rel=1e-12)
    assert welch["t"] == pytest.approx(expected_welch["welch_t"], rel=1e-12)
    assert welch["df"] == pytest.approx(expected_welch["welch_df"], rel=1e-12)
    assert welch["p"] == pytest.approx(expected_welch["p_two_sided"], rel=1e-12)
