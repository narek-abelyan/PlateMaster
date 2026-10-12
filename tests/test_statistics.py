import math

import numpy as np
import pandas as pd
import pytest
from scipy import stats


def metrics(frame):
    return frame.set_index("Metric")["Value"]


def test_reviewer_welch_example(app_module):
    result = metrics(app_module._welch_t_test([1, 2], [4, 5]))
    assert result["Welch t"] == pytest.approx(-4.242640687119285)
    assert result["df (Welch)"] == pytest.approx(2)
    # Independent analytical result, not the previous normal-tail approximation.
    assert result["p-value (2-sided, Welch t)"] == pytest.approx(0.0513167019494862)
    assert result["p-value (2-sided, Welch t)"] > 0.05
    assert app_module._welch_p_value([1, 2], [4, 5]) == result["p-value (2-sided, Welch t)"]


@pytest.mark.parametrize("a,b", [
    ([1, 2, 4], [0, 3, 7, 11]),
    ([1, 1, 1], [1, 2, 3]),
    ([0.01, 0.02, 0.07, 0.09], [0.04, 0.12, 0.8]),
])
def test_welch_matches_reference_for_unequal_samples(app_module, a, b):
    expected = stats.ttest_ind(a, b, equal_var=False)
    result = app_module._welch_result(a, b)
    assert result["p"] == pytest.approx(expected.pvalue)
    assert result["df"] == pytest.approx(expected.df)
    assert result["t"] == pytest.approx(expected.statistic)


def test_mean_ci_and_one_sample_test(app_module):
    result = metrics(app_module._basic_stat_metrics([1, 2]))
    # For n=2, SEM=0.5 and t_(0.975,1)=12.706204736.
    half_width = 6.353102368087347
    assert result["95% CI half-width (mean, Student t)"] == pytest.approx(half_width)
    assert result["95% CI lower (mean)"] == pytest.approx(1.5 - half_width)
    assert result["95% CI upper (mean)"] == pytest.approx(1.5 + half_width)
    # df=1 has a Cauchy distribution; t=3 gives this exact tail probability.
    expected_p = 1 - 2 * math.atan(3) / math.pi
    assert result["p-value (mean != 0, Student t)"] == pytest.approx(expected_p)


def test_missing_values_use_available_sample_size(app_module):
    result = metrics(app_module._basic_stat_metrics([1, np.nan, 2]))
    assert result["n"] == 2
    assert result["mean"] == 1.5
    result = app_module._welch_result([1, np.nan, 2], [4, 5, np.nan])
    assert (result["n_a"], result["n_b"]) == (2, 2)
    assert result["p"] == pytest.approx(0.0513167019494862)


@pytest.mark.parametrize("a,b", [([], [1, 2]), ([1], [2, 3]), ([1, 1], [1, 1]), ([1, 1], [2, 2])])
def test_undefined_welch_has_no_false_significance(app_module, a, b):
    result = app_module._welch_result(a, b)
    assert math.isnan(result["p"])
    assert result["status"] != "OK"
    assert math.isnan(app_module._welch_p_value(a, b))


def test_singleton_and_constant_sample_policy(app_module):
    one = metrics(app_module._basic_stat_metrics([7]))
    assert math.isnan(one["sd"])
    assert math.isnan(one["sem"])
    assert math.isnan(one["95% CI half-width (mean, Student t)"])
    constant = metrics(app_module._basic_stat_metrics([7, 7, 7]))
    assert math.isnan(constant["p-value (mean != 0, Student t)"])
    assert "zero sample variance" in constant["One-sample test status"]
    assert constant["95% CI lower (mean)"] == 7
    assert metrics(app_module._basic_stat_metrics([np.nan]))["No data"] == "N/A"


def test_infinite_input_is_rejected(app_module):
    with pytest.raises(ValueError, match="infinite"):
        app_module._welch_p_value([1, float("inf")], [2, 3])


@pytest.mark.parametrize("layout", ["columns", "rows", "groups"])
def test_every_layout_uses_student_t_uncertainty(app_module, layout):
    m = app_module
    if layout == "columns":
        table = pd.DataFrame({"row_id": ["A", "B"], "1": [1., 2.]})
        result = m.aggregate_rows(table, ["A", "B"], ["1"], m.AGGREGATIONS["Mean"])
        key = "1"
    elif layout == "rows":
        table = pd.DataFrame({"row_id": ["A"], "1": [1.], "2": [2.]})
        result = m.aggregate_columns_by_row(table, ["A"], ["1", "2"], m.AGGREGATIONS["Mean"])
        key = "A"
    else:
        table = pd.DataFrame({"row_id": ["A", "B"], "1": [1., 2.]})
        result = m.aggregate_by_groups(table, {"control": [("A", "1"), ("B", "1")]}, m.AGGREGATIONS["Mean"])
        key = "control"
    aggregated, _, errors, _ = result
    assert aggregated.iloc[0][key] == 1.5
    assert errors[key]["ci95"] == pytest.approx(6.353102368087347)


def test_removed_aggregation_is_not_offered(app_module):
    assert set(app_module.AGGREGATIONS) == {"Mean", "Median", "Min", "Max", "Std", "Sum"}
