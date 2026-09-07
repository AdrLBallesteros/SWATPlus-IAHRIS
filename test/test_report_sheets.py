"""Slicing the IAHRIS master workbook into the seven thematic reports."""

import os

import pytest

from conftest import master_workbook, sheet_names

# IAHRIS does not emit the same sheets for both published runs: the set depends
# on the flow regime it classifies the series into.
REFERENCE_SHEETS = set(sheet_names(master_workbook("CC"))) | set(
    sheet_names(master_workbook("DAM"))
)

# Checkbox, tick label and file name of each theme of the report window.
THEMES = [
    ("checkBox_nat", "label_nat", "SWATPlus-IAHRIS_Natural_Flow_Characterization.xlsx"),
    ("checkBox_alt", "label_alt", "SWATPlus-IAHRIS_Altered_Flow_Characterization.xlsx"),
    (
        "checkBox_nat_alt",
        "label_nat_alt",
        "SWATPlus-IAHRIS_Natural-Altered_Flow_Comparison.xlsx",
    ),
    ("checkBox_curves", "label_curves", "SWATPlus-IAHRIS_Flow_Rates_Duration_Curves.xlsx"),
    ("checkBox_habitual", "label_habitual", "SWATPlus-IAHRIS_IHA_Habitual_Values.xlsx"),
    ("checkBox_floods", "label_floods", "SWATPlus-IAHRIS_IHA_Floods_Droughts.xlsx"),
    (
        "checkBox_sign",
        "label_sign",
        "SWATPlus-IAHRIS_IHA_Environmental_Significance.xlsx",
    ),
]

# 'Informe nº8a' and 'Informe nº8b' are the same report for two different flow
# regimes and are handled asymmetrically, see the known limitations in
# test/README.md.  The consistency test below skips this pair.
KNOWN_GAPS = {"Informe nº8a", "Informe nº8b"}


@pytest.fixture
def reports_window(main_module, qt_app, tmp_path, monkeypatch):
    monkeypatch.setattr(os, "startfile", lambda path: None)

    window = main_module.ReportsWindow(
        str(tmp_path / "master.xlsx"), str(tmp_path / "out")
    )
    window.default_selection = {
        checkbox: getattr(window, checkbox).isChecked() for checkbox, _, _ in THEMES
    }
    for checkbox, _, _ in THEMES:
        getattr(window, checkbox).setChecked(False)

    window.extracted = []
    window.renamed = []
    monkeypatch.setattr(
        window,
        "extract_selected_sheets_to_excel",
        lambda names, path: window.extracted.append((list(names), path)),
    )
    monkeypatch.setattr(
        window, "rename_sheets_in_excel", lambda path: window.renamed.append(path)
    )
    yield window
    window.close()


def extract_theme(reports_window, checkbox):
    getattr(reports_window, checkbox).setChecked(True)
    reports_window.on_print_button_clicked()
    assert len(reports_window.extracted) == 1
    return reports_window.extracted[0]


@pytest.mark.parametrize("checkbox, label, file_name", THEMES)
def test_theme_writes_its_own_report(reports_window, checkbox, label, file_name):
    sheets, path = extract_theme(reports_window, checkbox)

    assert os.path.basename(path) == file_name
    assert os.path.dirname(path) == reports_window.report_folder
    assert sheets, "a theme must extract at least one sheet"
    assert reports_window.renamed == [path], "the sheets must also be renamed"
    assert getattr(reports_window, label).text() == "✓"


def test_all_themes_are_selected_when_the_window_opens(reports_window):
    assert all(reports_window.default_selection.values())


def test_nothing_is_extracted_when_no_theme_is_selected(reports_window):
    reports_window.on_print_button_clicked()

    assert reports_window.extracted == []
    assert reports_window.renamed == []


def test_every_theme_can_be_printed_at_once(reports_window):
    for checkbox, _, _ in THEMES:
        getattr(reports_window, checkbox).setChecked(True)

    reports_window.on_print_button_clicked()

    produced = [os.path.basename(path) for _, path in reports_window.extracted]
    assert produced == [file_name for _, _, file_name in THEMES]
    assert len(set(produced)) == len(THEMES), "the file names must be distinct"


def test_themes_do_not_overlap(reports_window):
    for checkbox, _, _ in THEMES:
        getattr(reports_window, checkbox).setChecked(True)
    reports_window.on_print_button_clicked()

    requested = [sheet for sheets, _ in reports_window.extracted for sheet in sheets]
    assert len(requested) == len(set(requested))


@pytest.mark.parametrize("checkbox, label, file_name", THEMES)
def test_requested_sheets_exist_in_iahris_output(
    reports_window, checkbox, label, file_name
):
    """A name no run produces would be skipped silently."""
    sheets, _ = extract_theme(reports_window, checkbox)

    unknown = [sheet for sheet in sheets if sheet not in REFERENCE_SHEETS]
    assert unknown == [], "not produced by IAHRIS: {}".format(unknown)


@pytest.mark.parametrize("checkbox, label, file_name", THEMES)
def test_requested_sheets_are_renamed_to_english(
    reports_window, main_module, checkbox, label, file_name
):
    """A copied sheet missing from the rename table keeps its Spanish name."""
    sheets, _ = extract_theme(reports_window, checkbox)
    rename_dict = rename_table(main_module)

    missing = [
        sheet
        for sheet in sheets
        if sheet not in rename_dict and sheet not in KNOWN_GAPS
    ]
    assert missing == [], "no English name for: {}".format(missing)


def test_rename_table_only_holds_names_iahris_emits(main_module):
    """A typo in the rename table would leave the sheet name unchanged."""
    unknown = [name for name in rename_table(main_module) if name not in REFERENCE_SHEETS]

    assert unknown == [], "not produced by IAHRIS: {}".format(unknown)


def test_english_names_are_unique(main_module):
    english = list(rename_table(main_module).values())

    assert len(english) == len(set(english))


def rename_table(main_module):
    """Read 'rename_dict' back from the source: it is a local variable."""
    import ast
    import inspect
    import textwrap

    source = textwrap.dedent(
        inspect.getsource(main_module.ReportsWindow.rename_sheets_in_excel)
    )
    for node in ast.walk(ast.parse(source)):
        if isinstance(node, ast.Assign) and any(
            isinstance(target, ast.Name) and target.id == "rename_dict"
            for target in node.targets
        ):
            return ast.literal_eval(node.value)
    raise AssertionError("'rename_dict' not found in rename_sheets_in_excel")


@pytest.fixture(scope="session")
def excel(main_module):
    try:
        app = main_module.xlwings.App(visible=False)
    except Exception as error:
        pytest.skip("Excel is not available: {}".format(error))
    app.quit()
    return main_module.xlwings


@pytest.mark.excel
@pytest.mark.parametrize("case", ["CC", "DAM"])
def test_extracted_workbook_holds_the_selected_sheets(
    main_module, qt_app, excel, tmp_path, monkeypatch, case
):
    monkeypatch.setattr(os, "startfile", lambda path: None)
    output_folder = tmp_path / case
    output_folder.mkdir()

    window = main_module.ReportsWindow(master_workbook(case), str(output_folder))
    try:
        window.checkBox_nat.setChecked(True)
        window.on_print_button_clicked()
    finally:
        window.close()

    produced = output_folder / "SWATPlus-IAHRIS_Natural_Flow_Characterization.xlsx"
    assert produced.exists()

    names = sheet_names(str(produced))
    assert names == ["Report_n1", "Report_n2", "Report_n2a", "Report_n4"]
