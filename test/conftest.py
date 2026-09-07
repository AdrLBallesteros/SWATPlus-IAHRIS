"""Shared fixtures for the SWATPlus-IAHRIS test suite."""

import contextlib
import glob
import os
import random
import re
import shutil
import sqlite3
import subprocess
import sys
import zipfile
from types import SimpleNamespace

import pytest

# Must be set before PyQt6 is imported.
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(REPO_ROOT, "test", "data")
SCENARIOS_DIR = os.path.join(DATA_DIR, "swatplus", "Scenarios")
CSV_DIR = os.path.join(DATA_DIR, "csv")
SERIES_DIR = os.path.join(REPO_ROOT, "test", "Natural-Altered_flow_series")
REPORTS_DIR = os.path.join(REPO_ROOT, "test", "Reports")

CC_NATURAL_SCENARIO = "scenario_cal_historical_cc_1990-2010"
CC_NATURAL_PERIOD = (1990, 2010)
CC_ALTERED_SCENARIO = "scenario_cal_future_ssp585_cc_2080-2100"
CC_ALTERED_PERIOD = (2080, 2100)
DAM_NATURAL_SCENARIO = "scenario_cal_toolbox_daily_500sim"
DAM_NATURAL_PERIOD = (1990, 2020)
SCENARIO_SHORT_NAME = "scenario_cal"
EMPTY_SCENARIO = "scenario_without_channel_output"

PAPER_CHANNEL = "58"
HEADWATER_CHANNEL = "56"

GAUGE_CSV = os.path.join(CSV_DIR, "Aforo_Belena.csv")
GAUGE_NAME = "Aforo_Belena"
GAUGE_PERIOD = DAM_NATURAL_PERIOD

sys.path.insert(0, REPO_ROOT)
INSTALL_DIR = "C:\\SWATPlus-IAHRIS"
TEMP_DIR = os.path.join(INSTALL_DIR, "temp")


def published_series(case, name):
    return os.path.join(SERIES_DIR, case, name)


def read_published_series(case, name):
    with open(published_series(case, name), encoding="utf-8") as handle:
        return handle.read()


def master_workbook(case):
    matches = glob.glob(os.path.join(REPORTS_DIR, case, "IAHRIS_Report_*.xlsx"))
    assert len(matches) == 1, "expected one master workbook in {}".format(case)
    return matches[0]


def sheet_names(workbook_path):
    with zipfile.ZipFile(workbook_path) as archive:
        workbook_xml = archive.read("xl/workbook.xml").decode("utf-8")
    return re.findall(r'<sheet name="([^"]+)"', workbook_xml)


@pytest.fixture(scope="session", autouse=True)
def _working_directory():
    # 'resource_path' resolves GUI.ui against the working directory.
    previous = os.getcwd()
    os.chdir(REPO_ROOT)
    yield
    os.chdir(previous)


@pytest.fixture(scope="session")
def qt_app(_working_directory):
    from PyQt6 import QtWidgets

    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    yield app


@pytest.fixture(scope="session")
def main_module(qt_app):
    import main

    return main


@pytest.fixture
def window(main_module, qt_app):
    window = main_module.MainWindow()
    yield window
    window.close()


@pytest.fixture
def messages(monkeypatch, main_module):
    """Record the message boxes as '(level, title, text)' instead of showing them."""
    from PyQt6 import QtWidgets

    recorded = []

    def _record(level):
        def _show(parent, title, text, *args, **kwargs):
            recorded.append((level, title, text))
            return QtWidgets.QMessageBox.StandardButton.Ok

        return _show

    for level in ("warning", "critical", "information", "question"):
        monkeypatch.setattr(
            QtWidgets.QMessageBox, level, staticmethod(_record(level))
        )
    return recorded


@pytest.fixture
def report_folder(tmp_path, monkeypatch):
    from PyQt6 import QtWidgets

    folder = tmp_path / "reports"
    folder.mkdir()
    monkeypatch.setattr(
        QtWidgets.QFileDialog,
        "getExistingDirectory",
        staticmethod(lambda *args, **kwargs: str(folder).replace("\\", "/")),
    )
    return folder


@pytest.fixture
def install_dir():
    # Only a folder created here is removed again, so an existing installation
    # of the software is left untouched.
    created = not os.path.exists(INSTALL_DIR)
    if created:
        os.makedirs(INSTALL_DIR)
    yield INSTALL_DIR
    if os.path.exists(TEMP_DIR):
        shutil.rmtree(TEMP_DIR, ignore_errors=True)
    if created:
        shutil.rmtree(INSTALL_DIR, ignore_errors=True)


class FakeIahris:
    """Stand-in for the IAHRIS.exe run driven by 'generate_report.bat'."""

    def __init__(self, report_folder, capture_dir):
        self.report_folder = str(report_folder)
        self.capture_dir = str(capture_dir)

        self.returncode = 0
        self.stdout = "IAHRIS 4.0\nProceso finalizado.\n"
        self.stderr = ""
        self.reports = ["IAHRIS_Report_SCENARIO_CAL_SCENARIO_CAL_COEDNO.xlsx"]

        self.called = False
        self.bat_content = ""

    def run(self, args, **kwargs):
        self.called = True
        bat_path = args[0] if isinstance(args, (list, tuple)) else args
        with open(bat_path, encoding="utf-8") as handle:
            self.bat_content = handle.read()

        # The temp folder is deleted on success, so keep a copy of the inputs.
        shutil.copytree(
            os.path.dirname(bat_path), self.capture_dir, dirs_exist_ok=True
        )

        for name in self.reports:
            with open(os.path.join(self.report_folder, name), "wb") as handle:
                handle.write(b"stub workbook")

        return subprocess.CompletedProcess(
            args, self.returncode, self.stdout, self.stderr
        )

    def input_file(self, name):
        with open(os.path.join(self.capture_dir, name), encoding="utf-8") as handle:
            return handle.read()

    def input_rows(self, name):
        content = self.input_file(name)
        return [line.split(";") for line in content.splitlines()]


@pytest.fixture
def fake_iahris(monkeypatch, main_module, report_folder, tmp_path):
    iahris = FakeIahris(report_folder, tmp_path / "captured_temp")
    # Replace only the 'subprocess' name seen by main.py.
    shim = SimpleNamespace(
        run=iahris.run, CREATE_NO_WINDOW=subprocess.CREATE_NO_WINDOW
    )
    monkeypatch.setattr(main_module, "subprocess", shim)
    return iahris


class _FakeCell:
    def __init__(self, store, key):
        self._store = store
        self._key = key

    @property
    def value(self):
        return self._store.get(self._key)

    @value.setter
    def value(self, value):
        self._store[self._key] = value


class _FakeSheet:
    def __init__(self, name):
        self.name = name
        self.cells = {}

    def __getitem__(self, key):
        return _FakeCell(self.cells, key)


class _FakeBook:
    def __init__(self, path):
        self.path = path
        self.sheets = [_FakeSheet("REPORTS")]
        self.saved = False
        self.closed = False

    def save(self, path=None):
        self.saved = True

    def close(self):
        self.closed = True


class FakeExcel:
    """Minimal xlwings.App replacement."""

    instances = []

    def __init__(self, visible=False):
        self.books = self
        self.opened = []
        self.quit_called = False
        FakeExcel.instances.append(self)

    def open(self, path):
        book = _FakeBook(path)
        self.opened.append(book)
        return book

    def add(self):
        return _FakeBook(None)

    def quit(self):
        self.quit_called = True


@pytest.fixture
def fake_excel(monkeypatch, main_module):
    FakeExcel.instances = []
    monkeypatch.setattr(main_module.xlwings, "App", FakeExcel)
    return FakeExcel


def shuffled_scenario(scenario, destination):
    """Copy a scenario database rewriting its records in a random order."""
    source_path = os.path.join(
        SCENARIOS_DIR, scenario, "Results", "swatplus_output.sqlite"
    )
    target_path = os.path.join(
        str(destination), "Scenarios", scenario, "Results", "swatplus_output.sqlite"
    )
    os.makedirs(os.path.dirname(target_path), exist_ok=True)

    source = sqlite3.connect(source_path)
    ddl = source.execute(
        "SELECT sql FROM sqlite_master WHERE name = 'channel_sd_day'"
    ).fetchone()[0]
    # 'id' is left out: as the primary key it would restore the original order.
    columns = [
        description[0]
        for description in source.execute(
            "SELECT * FROM channel_sd_day LIMIT 1"
        ).description
        if description[0] != "id"
    ]
    rows = source.execute(
        "SELECT {} FROM channel_sd_day ORDER BY id".format(", ".join(columns))
    ).fetchall()
    source.close()

    random.Random(20250504).shuffle(rows)

    target = sqlite3.connect(target_path)
    target.execute(ddl)
    target.executemany(
        "INSERT INTO channel_sd_day ({}) VALUES ({})".format(
            ", ".join(columns), ", ".join("?" for _ in columns)
        ),
        rows,
    )
    target.commit()
    target.close()
    return os.path.join(str(destination), "Scenarios")


@contextlib.contextmanager
def patched_dialog(name, replacement):
    """Answer one QFileDialog call, then put the previous behaviour back."""
    from PyQt6 import QtWidgets

    missing = object()
    previous = QtWidgets.QFileDialog.__dict__.get(name, missing)
    setattr(QtWidgets.QFileDialog, name, staticmethod(replacement))
    try:
        yield
    finally:
        if previous is missing:
            delattr(QtWidgets.QFileDialog, name)
        else:
            setattr(QtWidgets.QFileDialog, name, previous)


def load_swatplus_natural(
    window, scenario=CC_NATURAL_SCENARIO, channel=PAPER_CHANNEL, scenarios_dir=None
):
    window.radioButton_swat_nat.setChecked(True)
    # A path already in the line edit takes the command line branch, no dialog.
    window.lineEdit_nat.setText(scenarios_dir or SCENARIOS_DIR)
    window.select_file_nat()
    window.comboBox_scenario_nat.setCurrentText(scenario)
    window.select_Scenario_nat()
    window.comboBox_channel_nat.setCurrentText(channel)


def load_swatplus_altered(
    window, scenario=CC_ALTERED_SCENARIO, channel=PAPER_CHANNEL, scenarios_dir=None
):
    folder = scenarios_dir or SCENARIOS_DIR
    with patched_dialog("getExistingDirectory", lambda *args, **kwargs: folder):
        window.radioButton_swat_alt.setChecked(True)
        window.select_file_alt()
    window.comboBox_scenario_alt.setCurrentText(scenario)
    window.select_Scenario_alt()
    window.comboBox_channel_alt.setCurrentText(channel)


def load_csv(window, path, side="nat"):
    chosen = (str(path), "CSV files (*.csv)")
    with patched_dialog("getOpenFileName", lambda *args, **kwargs: chosen):
        if side == "nat":
            window.radioButton_csv_nat.setChecked(True)
            window.select_file_nat()
        else:
            window.radioButton_csv_alt.setChecked(True)
            window.select_file_alt()


def set_period(window, start_year, end_year, side="nat"):
    from PyQt6.QtCore import QDate

    getattr(window, "DateEdit_start_year_" + side).setDate(QDate(start_year, 1, 1))
    getattr(window, "DateEdit_finish_year_" + side).setDate(QDate(end_year, 1, 1))
