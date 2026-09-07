# SWATPlus-IAHRIS test suite

Automated regression tests for the SWATPlus-IAHRIS workflow, run against the
data of the published Sorbe River case study. The suite runs unattended on
Windows, needs neither a display nor a licence of IAHRIS, and is executed on
every push through [GitHub Actions](../.github/workflows/tests.yml).

## Running the tests

```powershell
# From the repository root, with the software installed from pyproject.toml
pip install -e ".[test]"
pytest

# Only one area
pytest test/test_csv_validation.py

# Also run the tests that drive Microsoft Excel (deselected by default)
pytest -m excel
```

`pytest` reports 98 tests. The two additional tests marked `excel` need
Microsoft Excel installed and no workbook open; they are deselected by default
so the suite stays runnable on a machine or a CI runner without Excel.

## What is covered

| Area | Test module |
| --- | --- |
| SWAT+ SQLite extraction, channel and period selection | `test_swatplus_extraction.py` |
| Chronological ordering of the exported series | `test_swatplus_extraction.py` |
| Leap years and complete-year counts | `test_swatplus_extraction.py` |
| CSV validation: column names, missing records, negative flows, empty values | `test_csv_validation.py` |
| Zero-flow days accepted as valid data | `test_csv_validation.py` |
| 15-year minimum period, on both sides of the boundary | `test_iahris_input_contract.py` |
| Scenario-name truncation to 12 characters | `test_iahris_input_contract.py` |
| Layout of the IAHRIS bulk-load file and of the generated `.bat` | `test_iahris_input_contract.py` |
| Regression against the published bulk-load files, for all four input paths | `test_swatplus_extraction.py`, `test_iahris_input_contract.py` |
| IAHRIS execution failures and the execution log | `test_iahris_execution.py` |
| Identification of the report produced by the run | `test_iahris_execution.py` |
| Extraction and renaming of the thematic report sheets | `test_report_sheets.py` |

The application logic lives inside the two Qt windows of
[main.py](../main.py), so the tests drive those windows directly: the widgets
are filled in as a user would fill them, the methods under test are called, and
the files they produce are inspected. Qt runs with `QT_QPA_PLATFORM=offscreen`,
and the message boxes are captured instead of being shown, so a rejected input
is asserted through the message the user would have seen.

### Test doubles

Two external programs cannot run on a CI runner and are replaced (see
[conftest.py](conftest.py)):

- **IAHRIS.exe** is closed-source and cannot be redistributed. `fake_iahris`
  keeps the `generate_report.bat` file the software wrote and a copy of the
  input files it generated, and reproduces the return codes, console messages
  and output workbooks of a real run — including the case where IAHRIS reports
  an error on the console while still exiting with code 0.
- **Excel**, driven through xlwings, is replaced by `fake_excel`. The tests
  marked `excel` use the real thing against the master workbooks in
  [Reports/](Reports).

Everything between those two ends — reading SWAT+ output, validating the user
CSV, building the IAHRIS input files and command line, detecting a failed run,
identifying the report, and slicing it into themes — is exercised for real.

## The data

All the test data comes from the published case study of the Sorbe River basin.
Two comparisons were run for the paper:

| Case | Natural regime | Altered regime | Period |
| --- | --- | --- | --- |
| **CC** — climate change | SWAT+ `scenario_cal_historical_cc_1990-2010` | SWAT+ `scenario_cal_future_ssp585_cc_2080-2100` | 1990–2010 vs 2080–2100 |
| **DAM** — reservoir impact | SWAT+ `scenario_cal_toolbox_daily_500sim` | Bélena gauge record, read from CSV | 1990–2020 |

Both were taken at **channel 58**, the basin outlet. Between them the two cases
exercise all four input paths of the software: natural and altered from SWAT+,
and natural and altered from a user CSV file.

```text
test/
├── SWATplus_Scenarios/          the three full SWAT+ projects (4.2 GB, not in git)
├── data/                        the part of them the tests need (5 MB, in git)
│   ├── make_reference_dataset.py
│   ├── swatplus/Scenarios/          extracted SWAT+ output databases
│   └── csv/                         the gauge record and the invalid files
├── Natural-Altered_flow_series/ the IAHRIS bulk-load files of the two runs
└── Reports/                     the IAHRIS master workbooks of the two runs
```

### `data/` — the extracted inputs

The SWAT+ output databases are 250–370 MB each, far over what a repository can
hold, so [make_reference_dataset.py](data/make_reference_dataset.py) copies out
of them the records the software actually reads and writes them into `data/` as
small databases that keep the original SWAT+ Editor `channel_sd_day` schema.
Two channels are kept, the outlet (58) and a headwater channel (56) whose flows
are three orders of magnitude smaller, so the tests can check the channel
filter. The full period of each scenario is kept, and the flow values, dates,
channel numbers and scenario folder names are the ones of the published runs.
The columns the software reads are copied verbatim; the remaining 59 columns of
the schema are left empty, which is what keeps the files small.

`data/swatplus/Scenarios/` also holds `scenario_without_channel_output`, the
same schema with no record in it: what SWAT+ Editor produces when *Daily >
Model Components > Channel* is left unchecked.

`data/csv/Aforo_Belena.csv` is the observed Bélena record in the input format
documented in the user manual (`Date` in mm/dd/yyyy, `Flow`), and the four
`invalid_*.csv` files are its first 40 days with one defect introduced in each,
so even the malformed fixtures carry observed values.

Rebuilding `data/` needs the full SWAT+ projects, which are not distributed:

```powershell
python test/data/make_reference_dataset.py
```

### `Natural-Altered_flow_series/` — the expected outputs

These are the IAHRIS bulk-load files of the two published runs. They are the
reference the tests compare against: running the software on the extracted
databases reproduces all four of them **byte for byte**, which is what ties the
suite to the results of the paper.

### `Reports/` — the IAHRIS master workbooks

The workbooks of the two published runs, together with the reports derived from
them. They are the reference for the sheet names IAHRIS actually emits: the set
is not the same for both runs, because it depends on the flow regime IAHRIS
classifies the series into.

