# PlateMaster documented validation example

This directory contains a small, entirely synthetic validation dataset with known inputs, a fixed plate map, complete analysis settings, expected numerical results, and a regression test. It is intended to verify the PlateMaster workflow rather than represent a biological experiment.

## Dataset

The `data` directory contains exactly five Excel files for 0, 5, 10, 15 and 20 minutes. Each file has one worksheet named `Plate`, with the header in row 1 and a 4 × 4 measurement table. The last number in each filename is the time in minutes.

The 16 wells are assigned as follows:

- `Control`: columns 1 and 2, rows A–D (8 observations per time point);
- `Treated`: columns 3 and 4, rows A–D (8 observations per time point).

The same mapping is provided in `plate_map.csv`. The values are deterministic and contain no missing values or outliers.

## Exact PlateMaster workflow

1. Start PlateMaster with `streamlit run PlateMaster1.2.py`.
2. Load all five files from `example_2/data`, either using the server-folder option or by uploading them together.
3. Keep the default import choices: worksheet 1, header row 1, `Auto / partial table`, and custom labels disabled.
4. Upload `example_2/analysis_settings.json` under **Import analysis settings (.json)** and select **Apply imported settings**.
5. Confirm that the layout is **By plate groups (template)**, aggregation is **Mean**, outlier filtering is **None**, and error bars are **95% CI**.
6. Press **Run**.
7. Compare the aggregated `Control` and `Treated` columns with `expected_group_results.csv`.
8. Compare the configured Welch test (`Control` at 0 minutes versus `Treated` at 20 minutes) with `expected_welch_comparison.csv`.

The JSON records the five input-file hashes, plate map and complete analysis settings. A software-environment warning can appear when it is opened with a different PlateMaster or dependency version; this is provenance information and does not prevent applying the settings.

## Expected results

`expected_group_results.csv` gives, for every time and group, `n`, mean, median, sample SD (`ddof=1`), SEM, the two-sided 95% Student-t confidence interval for the mean, minimum, maximum, and the two-sided one-sample t-test against zero.

`expected_welch_comparison.csv` gives the sample sizes, means, Welch statistic, Welch–Satterthwaite degrees of freedom, and two-sided p-value for the configured comparison. Values are stored with up to 15 significant digits. Tests should use a floating-point tolerance rather than compare formatted strings.

Run only the validation-example regression test with:

```bash
python -m pytest -q tests/test_validation_example.py
```

Regenerate the five workbooks and reference files from the transparent source definition with:

```bash
python example_2/generate_dataset.py
```

## Files and license

- `data/*.xlsx`: five synthetic input workbooks;
- `plate_map.csv`: known well-to-group mapping;
- `analysis_settings.json`: reloadable PlateMaster configuration and input hashes;
- `expected_group_results.csv`: expected descriptive and one-sample statistics;
- `expected_welch_comparison.csv`: expected two-group result;
- `manifest.json`: file hashes, time unit and dataset license;
- `generate_dataset.py`: deterministic data and reference-file generator;
- `LICENSE.txt`: CC0 1.0 dedication for this validation dataset.

The files in `example_2` are released under CC0 1.0. They contain no human, animal or confidential measurements.
