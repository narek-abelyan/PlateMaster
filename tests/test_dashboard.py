from copy import deepcopy
from io import BytesIO
from pathlib import Path

import pandas as pd
import pytest
import streamlit as st
from streamlit.testing.v1 import AppTest

from analysis_settings import dumps_config, loads_config

APP = Path(__file__).resolve().parents[1] / "PlateMaster1.2.py"


def click(at, label):
    next(x for x in at.button if x.label == label).click().run()
    assert not at.exception, [x.message for x in at.exception]
    return at


@pytest.fixture
def spreadsheets(tmp_path):
    for time in (0, 10):
        pd.DataFrame({"Well": ["A", "B", "C"], "1": [1 + time, 2 + time, 3 + time],
                      "2": [4 + time, 5 + time, 6 + time]}).to_excel(tmp_path / f"plate_{time}.xlsx", index=False)
    return tmp_path


def load_folder(folder):
    at = AppTest.from_file(str(APP), default_timeout=30).run()
    at.radio[0].set_value("Server folder path").run()
    next(x for x in at.text_input if x.label == "Server folder path with Excel files").set_value(str(folder)).run()
    click(at, "Load folder")
    return at


@pytest.fixture
def downloads(monkeypatch):
    captured = {}
    original = st.download_button

    def record(label, data, *args, **kwargs):
        captured[label] = data
        return original(label, data, *args, **kwargs)

    monkeypatch.setattr(st, "download_button", record)
    return captured


def imported_settings(at, monkeypatch, payload):
    original = st.file_uploader

    def uploaded(label, *args, **kwargs):
        if kwargs.get("key") == "settings_upload":
            # AppTest has no file-uploader interaction: simulate the received file bytes.
            buffer = BytesIO(payload)
            st.session_state["settings_upload"] = buffer
            return buffer
        return original(label, *args, **kwargs)

    with monkeypatch.context() as context:
        context.setattr(st, "file_uploader", uploaded)
        at.run()
        click(at, "Apply imported settings")
    return at


def test_named_preset_applies_before_widgets(spreadsheets):
    at = load_folder(spreadsheets)
    assert "None" not in at.selectbox(key="agg_name").options
    assert "None" in at.selectbox(key="outlier_method").options
    click(at, "Run")
    at.selectbox(key="stat_time").set_value(10.).run()
    at.selectbox(key="stat_metric").set_value("2").run()
    at.selectbox(key="error_method").set_value("95% CI").run()
    at.text_input(key="preset_name").set_value("saved mean").run()
    click(at, "Save preset now")
    at.selectbox(key="agg_name").set_value("Median").run()
    assert "95% CI" not in at.selectbox(key="error_method").options
    at.selectbox(key="preset_to_load").set_value("saved mean").run()
    click(at, "Apply preset")
    assert at.selectbox(key="agg_name").value == "Mean"
    click(at, "Run")
    assert at.selectbox(key="stat_time").value == 10.
    assert at.selectbox(key="stat_metric").value == "2"
    assert at.selectbox(key="error_method").value == "95% CI"


def test_json_replay_in_new_session_preserves_group_results(spreadsheets, monkeypatch, downloads):
    first = load_folder(spreadsheets)
    config = deepcopy(first.session_state["_current_analysis_config"])
    config["settings"].update(agg_layout_mode="By plate groups (template)", plot_columns=["control", "treated"],
                              stat_time=10., stat_metric="treated", ttest_time_a=0., ttest_time_b=10.,
                              ttest_metric_a="control", ttest_metric_b="treated", error_method="95% CI",
                              outlier_method="IQR", iqr_multiplier=2.0)
    config["plate_groups"] = [{"row_id": row, "column": col, "group": group}
                               for row in ("A", "B", "C") for col, group in (("1", "control"), ("2", "treated"))]
    imported_settings(first, monkeypatch, dumps_config(config))
    click(first, "Run")
    expected = pd.read_excel(BytesIO(downloads["Export report.xlsx"]), sheet_name=None)
    snapshot = downloads["Download settings.json"]
    assert expected["Aggregated table"]["control"].tolist() == [2, 12]
    assert expected["Aggregated table"]["treated"].tolist() == [5, 15]
    second = load_folder(spreadsheets)
    imported_settings(second, monkeypatch, snapshot)
    assert not second.exception
    assert second.session_state["_current_analysis_config"]["plate_groups"] == config["plate_groups"]
    click(second, "Run")
    actual = pd.read_excel(BytesIO(downloads["Export report.xlsx"]), sheet_name=None)
    for sheet in ("Aggregated table", "Statistical analysis", "T-test analysis", "Plot data", "Plate groups", "Input files"):
        pd.testing.assert_frame_equal(expected[sheet], actual[sheet])
    assert loads_config(downloads["Download settings.json"])["settings"] == loads_config(snapshot)["settings"]
    assert any("match" in x.value for x in second.info)
    assert not any("not an exact replay" in x.value for x in second.warning)


def test_changed_file_is_detected(spreadsheets, monkeypatch):
    first = load_folder(spreadsheets)
    payload = dumps_config(first.session_state["_current_analysis_config"])
    pd.DataFrame({"Well": ["A", "B", "C"], "1": [100, 200, 300], "2": [4, 5, 6]}).to_excel(spreadsheets / "plate_0.xlsx", index=False)
    second = load_folder(spreadsheets)
    imported_settings(second, monkeypatch, payload)
    warnings = "\n".join(x.value for x in second.warning)
    assert "Input file changed: plate_0.xlsx" in warnings
    assert "not an exact replay" in warnings


def test_invalid_import_leaves_existing_choices(spreadsheets, monkeypatch):
    at = load_folder(spreadsheets)
    at.selectbox(key="agg_name").set_value("Median").run()
    broken = deepcopy(at.session_state["_current_analysis_config"])
    broken["settings"]["agg_name"] = "None"
    imported_settings(at, monkeypatch, dumps_config(broken))
    assert at.selectbox(key="agg_name").value == "Median"
    assert any("agg_name" in x.value for x in at.error)


def test_bundled_xls_files_have_original_byte_checksums(app_module):
    folder = APP.parent / "example"
    at = load_folder(folder)
    manifest = at.session_state["_current_analysis_config"]["inputs"]
    assert len(manifest) == 9
    assert all(len(x["sha256"]) == 64 for x in manifest)
    # Limit the comparison matrix while checking the actual report pipeline.
    at.multiselect(key="selected_columns").set_value(["1", "2"]).run()
    click(at, "Run")
    assert any(x.label == "Export report.xlsx" for x in at.get("download_button"))


def test_filter_thresholds_survive_hidden_widgets(spreadsheets):
    at = load_folder(spreadsheets)
    at.selectbox(key="outlier_method").set_value("IQR").run()
    at.slider(key="iqr_multiplier").set_value(2.4).run()
    at.selectbox(key="outlier_method").set_value("MAD").run()
    at.slider(key="mad_threshold").set_value(4.8).run()
    at.selectbox(key="outlier_method").set_value("None").run()
    at.selectbox(key="outlier_method").set_value("IQR").run()
    assert not at.exception
    assert at.slider(key="iqr_multiplier").value == 2.4
    settings = at.session_state["_current_analysis_config"]["settings"]
    assert settings["mad_threshold"] == 4.8
    at.selectbox(key="outlier_method").set_value("MAD").run()
    assert at.slider(key="mad_threshold").value == 4.8


def test_upload_and_zip_capture_original_file_hashes(monkeypatch):
    import hashlib
    import zipfile

    class Upload(BytesIO):
        def __init__(self, name, content):
            super().__init__(content)
            self.name = name

    output = BytesIO()
    pd.DataFrame({"Well": ["A", "B"], "1": [1, 2], "2": [4, 5]}).to_excel(output, index=False)
    content = output.getvalue()
    archive = BytesIO()
    with zipfile.ZipFile(archive, "w") as z:
        z.writestr("nested/plate_10.xlsx", content)
    original = st.file_uploader

    def uploads(label, *args, **kwargs):
        if label == "Upload Excel files from your local PC":
            return [Upload("plate_0.xlsx", content)]
        if label == "Or upload ONE ZIP with many Excel files":
            return Upload("plates.zip", archive.getvalue())
        return original(label, *args, **kwargs)

    monkeypatch.setattr(st, "file_uploader", uploads)
    at = AppTest.from_file(str(APP), default_timeout=30).run()
    click(at, "Load uploaded files")
    manifest = at.session_state["_current_analysis_config"]["inputs"]
    assert [x["file"] for x in manifest] == ["plate_0.xlsx", "plate_10.xlsx"]
    assert all(x["sha256"] == hashlib.sha256(content).hexdigest() for x in manifest)
    assert [x["time"] for x in manifest] == [0., 10.]
    click(at, "Run")
    assert any(x.label == "Export report.xlsx" for x in at.get("download_button"))


def test_missing_selection_does_not_partially_apply_config(spreadsheets, monkeypatch):
    at = load_folder(spreadsheets)
    config = deepcopy(at.session_state["_current_analysis_config"])
    config["settings"].update(agg_name="Sum", selected_rows=["not-in-the-data"])
    imported_settings(at, monkeypatch, dumps_config(config))
    assert at.selectbox(key="agg_name").value == "Mean"
    assert any("missing" in x.value for x in at.error)


def test_analysis_failure_names_file_and_blocks_partial_report(spreadsheets, downloads):
    # Exercise a real file-specific aggregation failure after successful import.
    at = load_folder(spreadsheets)
    broken = at.session_state["loaded_tables"]
    broken["plate_10.xlsx"] = broken["plate_10.xlsx"].drop(columns="2")
    at.session_state["loaded_tables"] = broken
    at.selectbox(key="agg_layout_mode").set_value("By original rows").run()
    click(at, "Run")
    assert any("plate_10.xlsx" in x.value and "Analysis stopped" in x.value for x in at.error)
    assert "Export report.xlsx" not in downloads


def test_import_controls_and_each_file_preview(tmp_path):
    output = tmp_path / "plate_0.xlsx"
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame({"Info": ["Actual Temperature:"]*2, "Value": [37, 37]}).to_excel(writer, sheet_name="Metadata", index=False)
        pd.DataFrame({"Well": ["A", "B"], "1": [1, 2]}).to_excel(writer, sheet_name="Readings", startrow=2, index=False)
    at = load_folder(tmp_path)
    assert any("Duplicate row identifiers" in x.value for x in at.warning)
    next(x for x in at.text_input if x.label.startswith("Worksheet name")).set_value("Readings").run()
    next(x for x in at.number_input if x.label.startswith("Header row")).set_value(3).run()
    click(at, "Load folder")
    assert at.session_state["loaded_tables"]["plate_0.xlsx"]["1"].tolist() == [1, 2]
    assert at.selectbox(key="preview_file").options == ["plate_0.xlsx"]
    pd.DataFrame({"Well": ["A", "B"], "1": [3, 4]}).to_excel(tmp_path / "plate_10.xlsx", sheet_name="Readings", startrow=2, index=False)
    click(at, "Load folder")
    at.selectbox(key="preview_file").set_value("plate_10.xlsx").run()
    assert not at.exception
    assert at.dataframe[0].value["1"].tolist() == [3, 4]
