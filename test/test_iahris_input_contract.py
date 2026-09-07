"""The contract of the files handed over to IAHRIS."""

from conftest import (
    CC_ALTERED_PERIOD,
    CC_ALTERED_SCENARIO,
    CC_NATURAL_PERIOD,
    CC_NATURAL_SCENARIO,
    DAM_NATURAL_PERIOD,
    DAM_NATURAL_SCENARIO,
    GAUGE_CSV,
    GAUGE_NAME,
    SCENARIO_SHORT_NAME,
    load_csv,
    load_swatplus_altered,
    load_swatplus_natural,
    read_published_series,
    set_period,
)

import pytest

SIDE_START = {"nat": CC_NATURAL_PERIOD[0], "alt": CC_ALTERED_PERIOD[0]}


@pytest.mark.parametrize("side", ["nat", "alt"])
@pytest.mark.parametrize(
    "years, accepted",
    [
        (13, False),
        (14, False),
        (15, True),
        (16, True),
    ],
)
def test_period_length_boundary(
    window,
    messages,
    install_dir,
    report_folder,
    fake_iahris,
    fake_excel,
    side,
    years,
    accepted,
):
    """IAHRIS needs at least 15 complete years in both periods."""
    load_swatplus_natural(window)
    load_swatplus_altered(window)
    set_period(window, *CC_NATURAL_PERIOD, side="nat")
    set_period(window, *CC_ALTERED_PERIOD, side="alt")

    start_year = SIDE_START[side]
    set_period(window, start_year, start_year + years - 1, side=side)

    window.generate_reports()

    if accepted:
        assert messages == []
        assert fake_iahris.called
    else:
        assert [title for _, title, _ in messages] == ["Invalid selected periods"]
        assert not fake_iahris.called, "IAHRIS must not be launched"


@pytest.fixture
def climate_change_run(
    window, messages, install_dir, report_folder, fake_iahris, fake_excel
):
    load_swatplus_natural(window)
    load_swatplus_altered(window)
    set_period(window, *CC_NATURAL_PERIOD, side="nat")
    set_period(window, *CC_ALTERED_PERIOD, side="alt")
    window.generate_reports()
    assert messages == []
    return fake_iahris


@pytest.fixture
def reservoir_run(
    window, messages, install_dir, report_folder, fake_iahris, fake_excel
):
    load_swatplus_natural(window, scenario=DAM_NATURAL_SCENARIO)
    load_csv(window, GAUGE_CSV, side="alt")
    set_period(window, *DAM_NATURAL_PERIOD, side="nat")
    window.generate_reports()
    assert messages == []
    return fake_iahris


def test_natural_header_record(climate_change_run):
    rows = climate_change_run.input_rows(CC_NATURAL_SCENARIO + "_nat.csv")

    assert rows[0] == ["DIARIO", "NATURAL", SCENARIO_SHORT_NAME]


def test_altered_header_record(climate_change_run):
    rows = climate_change_run.input_rows(CC_ALTERED_SCENARIO + "_alt.csv")

    assert rows[0] == ["DIARIO", "ALTERADO", SCENARIO_SHORT_NAME, SCENARIO_SHORT_NAME]


def test_scenario_names_are_truncated_to_twelve_characters(climate_change_run):
    """Both climate scenarios share their first 12 characters, so IAHRIS sees
    one and the same name."""
    assert len(CC_NATURAL_SCENARIO) > 12 and len(CC_ALTERED_SCENARIO) > 12
    assert CC_NATURAL_SCENARIO[:12] == CC_ALTERED_SCENARIO[:12] == SCENARIO_SHORT_NAME

    header = climate_change_run.input_rows(CC_NATURAL_SCENARIO + "_nat.csv")[0]
    assert header[2] == SCENARIO_SHORT_NAME

    bat = climate_change_run.bat_content
    assert '/np:"{}"'.format(SCENARIO_SHORT_NAME) in bat
    assert '/na:"{}"'.format(SCENARIO_SHORT_NAME) in bat
    assert '/np:"{}"'.format(CC_NATURAL_SCENARIO) not in bat
    assert '/na:"{}"'.format(CC_ALTERED_SCENARIO) not in bat


def test_gauge_file_name_becomes_the_altered_scenario_name(reservoir_run):
    rows = reservoir_run.input_rows(GAUGE_NAME + "_alt.csv")

    assert rows[0] == ["DIARIO", "ALTERADO", SCENARIO_SHORT_NAME, GAUGE_NAME]


def test_climate_change_case_matches_the_published_files(climate_change_run):
    assert climate_change_run.input_file(
        CC_NATURAL_SCENARIO + "_nat.csv"
    ) == read_published_series("CC", "historical_nat.csv")
    assert climate_change_run.input_file(
        CC_ALTERED_SCENARIO + "_alt.csv"
    ) == read_published_series("CC", "future_ssp585_alt.csv")


def test_reservoir_case_matches_the_published_files(reservoir_run):
    assert reservoir_run.input_file(
        DAM_NATURAL_SCENARIO + "_nat.csv"
    ) == read_published_series("DAM", "scenario_cal_nat.csv")
    assert reservoir_run.input_file(
        GAUGE_NAME + "_alt.csv"
    ) == read_published_series("DAM", "Aforo_Belena_alt.csv")


def test_a_csv_file_can_also_be_the_natural_series(
    window, messages, install_dir, report_folder, fake_iahris, fake_excel
):
    load_csv(window, GAUGE_CSV, side="nat")
    load_swatplus_altered(window)
    set_period(window, *CC_ALTERED_PERIOD, side="alt")

    window.generate_reports()

    assert messages == []
    rows = fake_iahris.input_rows(GAUGE_NAME + "_nat.csv")
    assert rows[0] == ["DIARIO", "NATURAL", GAUGE_NAME]

    published = [
        line.split(";")[:2]
        for line in read_published_series("DAM", "Aforo_Belena_alt.csv").splitlines()
    ]
    assert [row[:2] for row in rows[1:]] == published[1:]


def test_file_has_no_header_line_and_uses_semicolons(climate_change_run):
    content = climate_change_run.input_file(CC_NATURAL_SCENARIO + "_nat.csv")
    first_line = content.splitlines()[0]

    assert not first_line.startswith("Date")
    assert "," not in content, "',' would be read as a column separator"
    assert all(line.count(";") == 2 for line in content.splitlines())


def test_sentinel_record_closes_the_last_year(climate_change_run):
    """IAHRIS only counts the last year when the following day is present."""
    end_year = CC_NATURAL_PERIOD[1]
    rows = climate_change_run.input_rows(CC_NATURAL_SCENARIO + "_nat.csv")

    assert rows[-2][0] == "31/12/{}".format(end_year)
    assert rows[-1][0] == "01/01/{}".format(end_year + 1)
    assert rows[-1][1] == "0.00"


def test_no_sentinel_record_when_the_period_ends_before_31_december(
    window, messages, install_dir, report_folder, fake_iahris, fake_excel
):
    load_swatplus_natural(window)
    load_swatplus_altered(window)
    set_period(window, CC_NATURAL_PERIOD[0], CC_NATURAL_PERIOD[1] + 1, side="nat")
    set_period(window, *CC_ALTERED_PERIOD, side="alt")

    window.generate_reports()

    assert messages == []
    rows = fake_iahris.input_rows(CC_NATURAL_SCENARIO + "_nat.csv")
    assert rows[-1][0] == "31/12/{}".format(CC_NATURAL_PERIOD[1])


def test_user_csv_dates_are_converted_to_day_first(reservoir_run):
    with open(GAUGE_CSV, encoding="utf-8") as handle:
        handle.readline()
        assert handle.readline().split(",")[0] == "01/01/1990"
        assert handle.readline().split(",")[0] == "01/02/1990"

    rows = reservoir_run.input_rows(GAUGE_NAME + "_alt.csv")
    assert rows[1][0] == "01/01/1990"
    assert rows[2][0] == "02/01/1990"


def test_bat_file_drives_the_three_iahris_calls(climate_change_run, report_folder):
    bat = climate_change_run.bat_content

    assert "cd C:\\SWATPlus-IAHRIS\\IAHRIS4.0" in bat
    assert "chcp 65001" in bat
    assert bat.count("IAHRIS.exe CD") == 2
    assert bat.count("IAHRIS.exe GIS") == 1
    assert "/t:P" in bat and "/t:A" in bat
    assert '/fs:"{}"'.format(report_folder) in bat
    assert bat.count("if errorlevel 1 exit /b %errorlevel%") == 4


def test_bat_file_points_at_the_generated_input_files(reservoir_run):
    bat = reservoir_run.bat_content

    assert (
        '/fe:"C:\\SWATPlus-IAHRIS\\temp\\{}_nat.csv"'.format(DAM_NATURAL_SCENARIO)
        in bat
    )
    assert '/fe:"C:\\SWATPlus-IAHRIS\\temp\\{}_alt.csv"'.format(GAUGE_NAME) in bat


def test_initial_month_of_the_water_year_is_passed_on(
    window, messages, install_dir, report_folder, fake_iahris, fake_excel
):
    load_swatplus_natural(window)
    load_swatplus_altered(window)
    set_period(window, *CC_NATURAL_PERIOD, side="nat")
    set_period(window, *CC_ALTERED_PERIOD, side="alt")
    window.comboBox_mon.setCurrentText("10")

    window.generate_reports()

    assert messages == []
    assert "/mi:10" in fake_iahris.bat_content
