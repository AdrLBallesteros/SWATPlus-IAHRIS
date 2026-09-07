"""SWAT+ SQLite extraction, channel selection and chronological ordering."""

import os
import sqlite3
from datetime import datetime

import pytest

from conftest import (
    CC_ALTERED_PERIOD,
    CC_ALTERED_SCENARIO,
    CC_NATURAL_PERIOD,
    CC_NATURAL_SCENARIO,
    DAM_NATURAL_PERIOD,
    DAM_NATURAL_SCENARIO,
    EMPTY_SCENARIO,
    HEADWATER_CHANNEL,
    PAPER_CHANNEL,
    SCENARIOS_DIR,
    load_swatplus_altered,
    load_swatplus_natural,
    read_published_series,
    set_period,
    shuffled_scenario,
)


def scenario_database(scenario):
    return os.path.join(SCENARIOS_DIR, scenario, "Results", "swatplus_output.sqlite")


def exported_dates(fake_iahris, name):
    rows = fake_iahris.input_rows(name)[1:]
    return [datetime.strptime(row[0], "%d/%m/%Y") for row in rows]


def test_scenarios_folder_populates_the_scenario_list(window, messages):
    load_swatplus_natural(window)

    offered = [
        window.comboBox_scenario_nat.itemText(index)
        for index in range(window.comboBox_scenario_nat.count())
    ]
    assert CC_NATURAL_SCENARIO in offered
    assert CC_ALTERED_SCENARIO in offered
    assert DAM_NATURAL_SCENARIO in offered
    assert messages == []


def test_a_folder_that_is_not_the_scenarios_folder_is_rejected(
    window, messages, tmp_path
):
    window.radioButton_swat_nat.setChecked(True)
    window.lineEdit_nat.setText(str(tmp_path))
    window.select_file_nat()

    assert [title for _, title, _ in messages] == ["Invalid Folder"]
    assert window.lineEdit_nat.text() == ""


@pytest.mark.parametrize(
    "scenario, period",
    [
        (CC_NATURAL_SCENARIO, CC_NATURAL_PERIOD),
        (CC_ALTERED_SCENARIO, CC_ALTERED_PERIOD),
        (DAM_NATURAL_SCENARIO, DAM_NATURAL_PERIOD),
    ],
)
def test_channels_and_period_are_read_from_the_database(
    window, messages, scenario, period
):
    load_swatplus_natural(window, scenario=scenario)

    channels = [
        window.comboBox_channel_nat.itemText(index)
        for index in range(window.comboBox_channel_nat.count())
    ]
    assert channels == [HEADWATER_CHANNEL, PAPER_CHANNEL]
    assert window.DateEdit_start_year_nat.date().year() == period[0]
    assert window.DateEdit_finish_year_nat.date().year() == period[1]
    assert messages == []


def test_scenario_without_channel_output_is_reported(window, messages):
    load_swatplus_natural(window, scenario=EMPTY_SCENARIO, channel="")

    assert window.comboBox_channel_nat.count() == 0
    assert [title for _, title, _ in messages] == ["Invalid SWAT+ Scenario"]


@pytest.fixture
def climate_change_run(
    window, messages, install_dir, report_folder, fake_iahris, fake_excel
):
    load_swatplus_natural(window)
    load_swatplus_altered(window)
    set_period(window, *CC_NATURAL_PERIOD, side="nat")
    set_period(window, *CC_ALTERED_PERIOD, side="alt")
    window.generate_reports()
    assert messages == [], "the run reported: {}".format(messages)
    return fake_iahris


def test_natural_series_matches_the_published_file(climate_change_run):
    generated = climate_change_run.input_file(CC_NATURAL_SCENARIO + "_nat.csv")

    assert generated == read_published_series("CC", "historical_nat.csv")


def test_altered_series_matches_the_published_file(climate_change_run):
    generated = climate_change_run.input_file(CC_ALTERED_SCENARIO + "_alt.csv")

    assert generated == read_published_series("CC", "future_ssp585_alt.csv")


def test_reservoir_case_matches_the_published_file(
    window, messages, install_dir, report_folder, fake_iahris, fake_excel
):
    load_swatplus_natural(window, scenario=DAM_NATURAL_SCENARIO)
    load_swatplus_altered(window)
    set_period(window, *DAM_NATURAL_PERIOD, side="nat")
    set_period(window, *CC_ALTERED_PERIOD, side="alt")

    window.generate_reports()

    assert messages == []
    generated = fake_iahris.input_file(DAM_NATURAL_SCENARIO + "_nat.csv")
    assert generated == read_published_series("DAM", "scenario_cal_nat.csv")


def test_series_is_chronological_even_if_the_database_is_not(
    window, messages, install_dir, report_folder, fake_iahris, fake_excel, tmp_path
):
    scrambled = shuffled_scenario(CC_NATURAL_SCENARIO, tmp_path / "scrambled")

    conn = sqlite3.connect(
        os.path.join(
            scrambled, CC_NATURAL_SCENARIO, "Results", "swatplus_output.sqlite"
        )
    )
    stored = conn.execute(
        "SELECT yr, mon, day FROM channel_sd_day LIMIT 400"
    ).fetchall()
    conn.close()
    assert stored != sorted(stored), "the copy is expected to be out of order"

    load_swatplus_natural(window, scenarios_dir=scrambled)
    load_swatplus_altered(window)
    set_period(window, *CC_NATURAL_PERIOD, side="nat")
    set_period(window, *CC_ALTERED_PERIOD, side="alt")
    window.generate_reports()
    assert messages == []

    generated = fake_iahris.input_file(CC_NATURAL_SCENARIO + "_nat.csv")
    assert generated == read_published_series("CC", "historical_nat.csv")

    dates = exported_dates(fake_iahris, CC_NATURAL_SCENARIO + "_nat.csv")
    assert all(
        (later - earlier).days == 1 for earlier, later in zip(dates, dates[1:])
    ), "the exported series must have no gaps and no repeated days"


def test_only_the_selected_channel_is_exported(
    window, messages, install_dir, report_folder, fake_iahris, fake_excel
):
    load_swatplus_natural(window, channel=HEADWATER_CHANNEL)
    load_swatplus_altered(window)
    set_period(window, *CC_NATURAL_PERIOD, side="nat")
    set_period(window, *CC_ALTERED_PERIOD, side="alt")
    window.generate_reports()
    assert messages == []

    rows = fake_iahris.input_rows(CC_NATURAL_SCENARIO + "_nat.csv")[1:-1]
    exported = [float(row[1]) for row in rows]

    conn = sqlite3.connect(scenario_database(CC_NATURAL_SCENARIO))
    expected = [
        row[0]
        for row in conn.execute(
            "SELECT flo_out FROM channel_sd_day WHERE unit = ? ORDER BY yr, mon, day",
            (int(HEADWATER_CHANNEL),),
        )
    ]
    conn.close()
    assert exported == expected
    assert exported != [
        float(row[1])
        for row in fake_iahris.input_rows(CC_ALTERED_SCENARIO + "_alt.csv")[1:-1]
    ]


def test_period_selection_filters_the_series(
    window, messages, install_dir, report_folder, fake_iahris, fake_excel
):
    load_swatplus_natural(window)
    load_swatplus_altered(window)
    set_period(window, 1995, 2010, side="nat")
    set_period(window, *CC_ALTERED_PERIOD, side="alt")
    window.generate_reports()
    assert messages == []

    dates = exported_dates(fake_iahris, CC_NATURAL_SCENARIO + "_nat.csv")
    assert dates[0] == datetime(1995, 1, 1)
    assert dates[-2] == datetime(2010, 12, 31)


@pytest.mark.parametrize("leap_year", [1992, 1996, 2000, 2004, 2008])
def test_leap_days_are_exported(climate_change_run, leap_year):
    rows = climate_change_run.input_rows(CC_NATURAL_SCENARIO + "_nat.csv")[1:]
    dates = {row[0] for row in rows}

    assert "29/02/{}".format(leap_year) in dates


@pytest.mark.parametrize("common_year", [1990, 1991, 1993, 1999, 2001, 2010])
def test_common_years_have_no_29_february(climate_change_run, common_year):
    rows = climate_change_run.input_rows(CC_NATURAL_SCENARIO + "_nat.csv")[1:]
    dates = {row[0] for row in rows}

    assert "29/02/{}".format(common_year) not in dates


def test_every_year_is_complete(climate_change_run):
    dates = exported_dates(climate_change_run, CC_NATURAL_SCENARIO + "_nat.csv")
    counted = {}
    for moment in dates:
        counted[moment.year] = counted.get(moment.year, 0) + 1

    start_year, end_year = CC_NATURAL_PERIOD
    for year in range(start_year, end_year + 1):
        leap = year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)
        assert counted[year] == (366 if leap else 365), "year {}".format(year)
    assert counted[end_year + 1] == 1, "the sentinel day of the following year"


def test_dates_use_the_day_first_format_iahris_requires(climate_change_run):
    start_year = CC_NATURAL_PERIOD[0]
    rows = climate_change_run.input_rows(CC_NATURAL_SCENARIO + "_nat.csv")[1:]

    assert rows[0][0] == "01/01/{}".format(start_year)
    assert rows[12][0] == "13/01/{}".format(start_year)
    for row in rows:
        day, month, year = row[0].split("/")
        assert len(day) == 2 and len(month) == 2 and len(year) == 4
