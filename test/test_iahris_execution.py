"""Driving IAHRIS.exe, detecting its failures and identifying its report."""

import os

import pytest

from conftest import (
    CC_ALTERED_PERIOD,
    CC_NATURAL_PERIOD,
    INSTALL_DIR,
    TEMP_DIR,
    load_swatplus_altered,
    load_swatplus_natural,
    set_period,
)


@pytest.fixture
def ready_window(window, install_dir, report_folder, fake_iahris, fake_excel):
    load_swatplus_natural(window)
    load_swatplus_altered(window)
    set_period(window, *CC_NATURAL_PERIOD, side="nat")
    set_period(window, *CC_ALTERED_PERIOD, side="alt")
    return window


def titles(messages):
    return [title for _, title, _ in messages]


def only_message(messages):
    assert len(messages) == 1, messages
    return messages[0]


def test_missing_installation_stops_the_run(
    ready_window, messages, fake_iahris, monkeypatch
):
    real_exists = os.path.exists
    monkeypatch.setattr(
        os.path,
        "exists",
        lambda path: False if path == INSTALL_DIR else real_exists(path),
    )

    ready_window.generate_reports()

    assert titles(messages) == ["Installation error"]
    assert not fake_iahris.called


def test_cancelled_report_folder_stops_the_run(
    ready_window, messages, fake_iahris, monkeypatch
):
    from PyQt6 import QtWidgets

    monkeypatch.setattr(
        QtWidgets.QFileDialog,
        "getExistingDirectory",
        staticmethod(lambda *args, **kwargs: ""),
    )

    ready_window.generate_reports()

    assert titles(messages) == ["No Folder Selected"]
    assert not fake_iahris.called


def test_non_zero_return_code_is_reported(ready_window, messages, fake_iahris):
    fake_iahris.returncode = 2
    fake_iahris.stdout = "Carga de datos interrumpida.\n"
    fake_iahris.reports = []

    ready_window.generate_reports()

    level, title, text = only_message(messages)
    assert (level, title) == ("critical", "IAHRIS Execution Error")
    assert "return code: 2" in text
    assert "Carga de datos interrumpida." in text


def test_error_message_on_a_zero_return_code_is_reported(
    ready_window, messages, fake_iahris
):
    """IAHRIS reports its own errors on the console and still exits with 0."""
    fake_iahris.returncode = 0
    fake_iahris.stdout = (
        "IAHRIS 4.0\n"
        "Error. El fichero de entrada no tiene el formato esperado.\n"
        "Proceso finalizado.\n"
    )
    fake_iahris.reports = []

    ready_window.generate_reports()

    level, title, text = only_message(messages)
    assert (level, title) == ("critical", "IAHRIS Execution Error")
    assert "reported an error" in text
    assert "Error. El fichero de entrada no tiene el formato esperado." in text


def test_a_run_without_any_report_is_reported(ready_window, messages, fake_iahris):
    fake_iahris.reports = []

    ready_window.generate_reports()

    level, title, text = only_message(messages)
    assert (level, title) == ("critical", "IAHRIS Execution Error")
    assert "no report file was generated" in text


def test_failed_run_keeps_the_inputs_and_writes_a_log(
    ready_window, messages, fake_iahris
):
    fake_iahris.returncode = 1
    fake_iahris.stderr = "Excepcion no controlada.\n"
    fake_iahris.reports = []

    ready_window.generate_reports()

    assert os.path.exists(TEMP_DIR), "the input files must be kept for diagnosis"
    with open(os.path.join(TEMP_DIR, "iahris_log.txt"), encoding="utf-8") as handle:
        log = handle.read()
    assert "Return code: 1" in log
    assert "Excepcion no controlada." in log
    assert "IAHRIS.exe GIS" in log

    _, _, text = only_message(messages)
    assert TEMP_DIR in text


def test_no_report_window_is_opened_after_a_failure(
    ready_window, messages, fake_iahris
):
    fake_iahris.returncode = 1
    fake_iahris.reports = []

    ready_window.generate_reports()

    assert not hasattr(ready_window, "reports_window")
    assert ready_window.progressBar.value() == 0


def test_workbook_already_in_the_folder_is_ignored(
    ready_window, messages, fake_iahris, report_folder
):
    older = report_folder / "IAHRIS_Report_FROM_LAST_WEEK.xlsx"
    older.write_bytes(b"older workbook")
    fake_iahris.reports = ["IAHRIS_Report_THIS_RUN.xlsx"]

    ready_window.generate_reports()

    assert messages == []
    assert ready_window.reports_window.last_generated_xlsx == str(
        report_folder / "IAHRIS_Report_THIS_RUN.xlsx"
    )


def test_several_new_workbooks_cannot_be_told_apart(
    ready_window, messages, fake_iahris
):
    fake_iahris.reports = ["IAHRIS_Report_A.xlsx", "IAHRIS_Report_B.xlsx"]

    ready_window.generate_reports()

    level, title, text = only_message(messages)
    assert (level, title) == ("critical", "IAHRIS Execution Error")
    assert "cannot be identified" in text
    assert "IAHRIS_Report_A.xlsx" in text and "IAHRIS_Report_B.xlsx" in text


def test_successful_run_cleans_up_and_opens_the_report_window(
    ready_window, messages, fake_iahris, report_folder
):
    ready_window.generate_reports()

    assert messages == []
    assert fake_iahris.called
    assert not os.path.exists(TEMP_DIR), "the temp folder is removed on success"

    reports_window = ready_window.reports_window
    assert reports_window.report_folder == str(report_folder)
    assert os.path.basename(reports_window.last_generated_xlsx) == fake_iahris.reports[0]


def test_successful_run_labels_the_flow_series_in_the_workbook(
    ready_window, messages, fake_excel
):
    ready_window.generate_reports()

    assert messages == []
    book = fake_excel.instances[0].opened[0]
    cells = book.sheets[0].cells
    assert cells["AA1"] == "Nat_F" and cells["AB1"] == "Natural Flow"
    assert cells["AA2"] == "Alt_F" and cells["AB2"] == "Altered Flow"
    assert cells["E3"] == "" and cells["E4"] == ""
    assert book.saved and book.closed


def test_workbook_that_cannot_be_opened_is_reported(
    ready_window, messages, monkeypatch, main_module
):
    """Excel refuses to open a workbook the user still has open."""

    def _raise(*args, **kwargs):
        raise OSError("the workbook is already open")

    monkeypatch.setattr(main_module.xlwings, "App", _raise)

    ready_window.generate_reports()

    assert titles(messages) == ["Excel File Access Error"]
    assert ready_window.progressBar.value() == 0
