from copy import deepcopy
from io import BytesIO
import json

import pandas as pd
import pytest

from analysis_settings import (
    DEFAULT_SETTINGS, MAX_SETTINGS_BYTES, dumps_config, input_record, loads_config,
    make_config, provenance_sheets, replay_warnings, validate_config,
)


@pytest.fixture
def config():
    settings = deepcopy(DEFAULT_SETTINGS)
    settings.update(agg_layout_mode="By plate groups (template)", selected_rows=["A", "B"],
                    selected_columns=["1"], plot_columns=["Контроль"],
                    stat_metric="Контроль", ttest_metric_a="Контроль", ttest_metric_b="Контроль")
    groups = [{"row_id": x, "column": "1", "group": "Контроль"} for x in ["A", "B"]]
    frame = pd.DataFrame({"row_id": ["A", "B"], "1": [1, 2]})
    return make_config(settings, groups, [input_record("plate_0.xlsx", b"original", frame, 0)],
                       {"application_version": "test", "python": "3.13", "libraries": {}, "source_sha256": {}})


def test_json_round_trip_includes_groups_and_inputs(config):
    loaded = loads_config(dumps_config(config))
    assert loaded == config
    loaded["plate_groups"][0]["group"] = "changed"
    assert config["plate_groups"][0]["group"] == "Контроль"


@pytest.mark.parametrize("key,value", [
    ("agg_name", "None"), ("outlier_method", "unknown"), ("iqr_multiplier", -1),
    ("mad_threshold", float("nan")), ("selected_rows", "AB"), ("stat_time", "yesterday"),
    ("selected_columns", ["1", "1"]), ("plot_columns", ["nonexistent"]),
])
def test_invalid_settings_are_rejected_without_mutation(config, key, value):
    broken = deepcopy(config)
    broken["settings"][key] = value
    with pytest.raises(ValueError):
        validate_config(broken)
    assert config["settings"] == loads_config(dumps_config(config))["settings"]


def test_unknown_schema_and_fields_are_rejected(config):
    for version in (2, True, None):
        broken = deepcopy(config)
        broken["schema_version"] = version
        with pytest.raises(ValueError, match="schema"):
            validate_config(broken)
    broken = deepcopy(config)
    broken["settings"]["loaded_tables"] = {}
    with pytest.raises(ValueError, match="unknown"):
        validate_config(broken)


def test_duplicate_group_cell_rejected(config):
    config["plate_groups"].append(deepcopy(config["plate_groups"][0]))
    with pytest.raises(ValueError, match="more than one group"):
        validate_config(config)


def test_wrong_mean_error_bars_rejected(config):
    config["settings"].update(agg_name="Sum", error_method="95% CI")
    with pytest.raises(ValueError, match="Mean"):
        validate_config(config)


def test_invalid_json_does_not_load():
    for content in (b"{", b"\xff", b"[]", b" " * (MAX_SETTINGS_BYTES + 1)):
        with pytest.raises(ValueError):
            loads_config(content)


def test_input_and_environment_changes_are_reported(config):
    assert replay_warnings(config, config["inputs"], config["software"]) == []
    changed = deepcopy(config["inputs"])
    changed[0]["sha256"] = "0" * 64
    assert any("Input file changed" in x for x in replay_warnings(config, changed, config["software"]))
    changed[0]["table_sha256"] = "0" * 64
    assert any("Parsed measurements changed" in x for x in replay_warnings(config, changed, config["software"]))
    software = deepcopy(config["software"])
    software["python"] = "different"
    assert any("python" in x for x in replay_warnings(config, config["inputs"], software))


def test_excel_contains_exact_reloadable_settings(app_module, config):
    # Long provenance must remain recoverable despite Excel's 32767-character cell limit.
    config["note"] = "x" * 65000
    report = app_module._build_excel_report(provenance_sheets(config))
    sheets = pd.read_excel(BytesIO(report), sheet_name=None)
    assert {"Analysis settings", "Plate groups", "Input files", "Software versions", "Settings JSON"} <= sheets.keys()
    parts = sheets["Settings JSON"].sort_values("Part")["JSON"]
    assert len(parts) > 1
    assert max(parts.str.len()) <= 30000
    assert loads_config("".join(parts).encode()) == config
    assert sheets["Plate groups"]["group"].tolist() == ["Контроль", "Контроль"]
    assert sheets["Input files"].iloc[0]["sha256"] == config["inputs"][0]["sha256"]


def test_pairwise_settings_old_files_and_selection_limit(config):
    old = deepcopy(config)
    old["settings"].pop("pairwise_enabled")
    old["settings"].pop("pairwise_groups")
    restored = validate_config(old)
    assert restored["settings"]["pairwise_enabled"] is False
    assert restored["settings"]["pairwise_groups"] == []
    config["settings"].update(pairwise_enabled=True, pairwise_groups=["0.0 | control", "10.0 | control"])
    assert loads_config(dumps_config(config)) == config
    config["settings"]["pairwise_groups"] = [str(i) for i in range(21)]
    with pytest.raises(ValueError, match="no more than 20"):
        validate_config(config)
