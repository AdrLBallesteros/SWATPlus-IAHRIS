"""Validation of the user-supplied CSV flow series."""

import csv
import os

import pytest

from conftest import CSV_DIR, GAUGE_CSV, GAUGE_PERIOD, load_csv

ZERO_FLOW_DAYS = [
    "11/12/1994",
    "11/13/1994",
    "11/14/1994",
    "11/15/1994",
    "11/16/1994",
    "11/17/1994",
    "06/26/1995",
]


def titles(messages):
    return [title for _, title, _ in messages]


def test_valid_series_is_accepted_and_sets_the_period(window, messages):
    load_csv(window, GAUGE_CSV)

    assert messages == []
    assert window.lineEdit_nat.text() == GAUGE_CSV
    assert window.DateEdit_start_year_nat.date().year() == GAUGE_PERIOD[0]
    assert window.DateEdit_finish_year_nat.date().year() == GAUGE_PERIOD[1]
    assert window.altered.isEnabled()


def test_zero_flow_days_are_valid_data(window, messages):
    """Zero flow is a value, not a gap: the gauge ran dry in November 1994."""
    with open(GAUGE_CSV, newline="", encoding="utf-8") as handle:
        flows = {row["Date"]: row["Flow"] for row in csv.DictReader(handle)}
    assert [day for day in ZERO_FLOW_DAYS if float(flows[day]) == 0] == ZERO_FLOW_DAYS

    load_csv(window, GAUGE_CSV)

    assert messages == []
    assert window.lineEdit_nat.text() == GAUGE_CSV


def test_a_cancelled_dialog_leaves_the_form_empty(window, messages):
    load_csv(window, "")

    assert messages == []
    assert window.lineEdit_nat.text() == ""


INVALID_FILES = [
    (
        "invalid_missing_columns.csv",
        "Invalid CSV Format",
        "columns are not named 'Date' and 'Flow'",
    ),
    (
        "invalid_missing_records.csv",
        "Invalid Date Frequency",
        "three days are missing in the middle of the series",
    ),
    (
        "invalid_negative_flow.csv",
        "Invalid Flow Data",
        "one flow value is negative",
    ),
    (
        "invalid_missing_flow_value.csv",
        "Invalid Flow Data",
        "one flow value is empty",
    ),
]


@pytest.mark.parametrize("side", ["nat", "alt"])
@pytest.mark.parametrize("name, title, description", INVALID_FILES)
def test_invalid_file_is_rejected(window, messages, side, name, title, description):
    if side == "alt":
        load_csv(window, GAUGE_CSV, side="nat")
        assert messages == []

    load_csv(window, os.path.join(CSV_DIR, name), side=side)

    assert titles(messages) == [title], description
    line_edit = window.lineEdit_nat if side == "nat" else window.lineEdit_alt
    assert line_edit.text() == "", "the rejected file must not stay selected"


def test_rejection_keeps_the_previously_selected_period(window, messages):
    load_csv(window, GAUGE_CSV, side="nat")
    load_csv(window, os.path.join(CSV_DIR, "invalid_negative_flow.csv"), side="alt")

    assert window.lineEdit_nat.text() == GAUGE_CSV
    assert window.DateEdit_start_year_nat.date().year() == GAUGE_PERIOD[0]
    assert window.DateEdit_finish_year_nat.date().year() == GAUGE_PERIOD[1]
