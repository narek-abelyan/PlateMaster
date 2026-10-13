"""Import checks use synthetic plate and instrument-metadata workbooks."""
from io import BytesIO

import numpy as np
import pandas as pd
import pytest


def workbook():
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame({"Instrument": ["Actual Temperature:"] * 9, "Value": [37.] * 9}).to_excel(
            writer, sheet_name="Metadata", index=False)
        pd.DataFrame({"Well": ["A", "B"], "1": [1, 2], "2": [4, 5]}).to_excel(
            writer, sheet_name="Readings", startrow=2, index=False)
    return output.getvalue()


def test_metadata_is_rejected_and_explicit_sheet_header_reads_measurements(app_module):
    data = workbook()
    with pytest.raises(ValueError, match="Duplicate row identifiers"):
        app_module.load_table_from_bytes(data, "PlateReader1.xlsx")
    table = app_module.load_table_from_bytes(data, "PlateReader1.xlsx", sheet_name="Readings", header=2)
    assert table["row_id"].tolist() == ["A", "B"]
    assert table["1"].tolist() == [1, 2]
    assert table.attrs["import_options"]["header_row"] == 3
    by_index = app_module.load_table_from_bytes(data, "PlateReader1.xlsx", sheet_name=1, header=2)
    pd.testing.assert_frame_equal(table, by_index)


def test_invalid_sheet_does_not_fall_back_to_first_sheet(app_module, monkeypatch):
    monkeypatch.setattr(app_module, "_try_libreoffice_convert_bytes_to_xlsx", lambda *args: None)
    with pytest.raises(ValueError, match="Unable to parse"):
        app_module.load_table_from_bytes(workbook(), "plate.xlsx", sheet_name="Missing")


@pytest.mark.parametrize("ids", [[None, "B"], ["A", "A"], [" A ", "A"], ["", "B"]])
def test_missing_and_duplicate_ids_are_rejected(app_module, ids):
    with pytest.raises(ValueError, match="identifiers"):
        app_module.sanitize_table(pd.DataFrame({"Well": ids, "1": [1, 2]}), allow_custom=True)


def test_custom_labels_require_opt_in_and_keep_warning(app_module):
    frame = pd.DataFrame({"Sample": ["control", "treated"], "signal": [1, 2]})
    with pytest.raises(ValueError, match="Non-standard"):
        app_module.sanitize_table(frame)
    table = app_module.sanitize_table(frame, allow_custom=True)
    assert "metadata" in table.attrs["import_warnings"][0]
    assert table["signal"].tolist() == [1, 2]


def test_expected_dimensions_and_well_coordinates(app_module):
    frame = pd.DataFrame({"Well": ["A", "B"], "1": [1, 2], "2": [3, 4], "3": [5, 6]})
    assert len(app_module.sanitize_table(frame, plate_shape=(2, 3))) == 2
    for invalid in (frame.iloc[:1], frame.rename(columns={"3": "4"}), frame.assign(Well=["A", "C"])):
        with pytest.raises(ValueError, match="Expected a 2 x 3"):
            app_module.sanitize_table(invalid, plate_shape=(2, 3))


def test_missing_values_are_reported_without_becoming_zero(app_module):
    frame = pd.DataFrame({"Well": ["A", "B", "C"], "1": [1, "bad", np.nan], "2": [2, 3, np.nan], "3": [None]*3})
    table = app_module.sanitize_table(frame)
    assert table["row_id"].tolist() == ["A", "B"]
    assert pd.isna(table.iloc[1]["1"])
    assert any("non-numeric" in w for w in table.attrs["import_warnings"])
    assert any("Removed 1 row(s) and 1 column(s)" in w for w in table.attrs["import_warnings"])


def test_infinite_measurements_are_rejected(app_module):
    with pytest.raises(ValueError, match="Infinite"):
        app_module.sanitize_table(pd.DataFrame({"Well": ["A"], "1": [np.inf]}))


def test_import_choices_are_exported_and_compared(app_module):
    from analysis_settings import input_record, make_config, DEFAULT_SETTINGS, provenance_sheets, replay_warnings
    from copy import deepcopy
    data = workbook()
    table = app_module.load_table_from_bytes(data, "plate_0.xlsx", sheet_name="Readings", header=2)
    record = input_record("plate_0.xlsx", data, table, 0)
    config = make_config(deepcopy(DEFAULT_SETTINGS), [], [record], {})
    sheets = provenance_sheets(config)
    assert sheets["Input files"].iloc[0]["import_options"]["sheet"] == "Readings"
    changed = deepcopy(record)
    changed["import_options"]["header_row"] = 4
    assert "Import choices changed: plate_0.xlsx" in replay_warnings(config, [changed], {})
