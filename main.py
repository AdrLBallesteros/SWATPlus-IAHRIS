"""
/***************************************************************************
 **SWATPlus-IAHRIS
 **A QGIS plugin
 **A software that links SWATPlus to IAHRIS to automatically generate reports on Indicators of Hydrologic Alteration in RIverS.
----------------------------------------------------
        begin                : **May-2025
        copyright            : **COPYRIGHT
        email                : **alopbal@upv.es
 ***************************************************************************/

/***************************************************************************
 *                                                                         *
 *   This program is free software; you can redistribute it and/or modify  *
 *   it under the terms of the GNU General Public License as published by  *
 *   the Free Software Foundation; either version 3 of the License, or     *
 *   any later version.                                                    *
 *                                                                         *
 ***************************************************************************/
"""

from PyQt6 import QtWidgets, uic
from PyQt6.QtCore import QDate
import os
import sqlite3
import pandas as pd
import subprocess
import shutil
from datetime import datetime
import xlwings
import glob


# https://stackoverflow.com/questions/7674790/bundling-data-files-with-pyinstaller-onefile/13790741#13790741
def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


# Define a MainWindow class that inherits from QMainWindow
class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        """Constructor of the SWATPlus-IAHRIS software - Main Window"""

        super().__init__()

        # Load the UI file
        uic.loadUi(resource_path("GUI.ui"), self)

        # Connect click events to a function
        self.pushButton_nat.clicked.connect(self.reset_select_file_nat)
        self.comboBox_scenario_nat.activated.connect(self.select_Scenario_nat)

        self.pushButton_alt.clicked.connect(self.select_file_alt)
        self.comboBox_scenario_alt.activated.connect(self.select_Scenario_alt)

        self.pushButton_reports.clicked.connect(self.generate_reports)

        self.radioButton_swat_nat.toggled.connect(self.reset_var_nat)
        self.radioButton_swat_alt.toggled.connect(self.reset_var_alt)

        # Deactivate SWAT+ parts of the GUI
        self.swatplus_nat.setEnabled(False)
        self.swatplus_alt.setEnabled(False)
        # Deactivate altered flow inputs
        self.altered.setEnabled(False)
        # Deactivate the reports button
        self.pushButton_reports.setEnabled(False)

    def reset_var_nat(self):
        """Reset variables when the nat radio button is toggled"""

        # Reset the folder and comboBox
        folder = None
        self.lineEdit_nat.setText(folder)
        self.comboBox_scenario_nat.clear()
        self.swatplus_nat.setEnabled(False)
        self.comboBox_channel_nat.clear()
        self.DateEdit_start_year_nat.setDate(QDate(2000, 1, 1))
        self.DateEdit_finish_year_nat.setDate(QDate(2000, 1, 1))

        # Reset SWAT+ parts of the GUI
        self.swatplus_nat.setEnabled(False)
        self.pushButton_nat.setEnabled(True)

    def reset_var_alt(self):
        """Reset variables when the alt radio button is toggled"""
        # Reset the folder and comboBox
        folder = None
        self.lineEdit_alt.setText(folder)
        self.comboBox_scenario_alt.clear()
        self.swatplus_alt.setEnabled(False)
        self.comboBox_channel_alt.clear()
        self.DateEdit_start_year_alt.setDate(QDate(2000, 1, 1))
        self.DateEdit_finish_year_alt.setDate(QDate(2000, 1, 1))

        # Reset the SWAT+ parts of the GUI and reports button
        self.pushButton_reports.setEnabled(False)
        self.swatplus_alt.setEnabled(False)

    def select_file_nat(self):
        """Open folder or file for nat input"""

        if self.radioButton_swat_nat.isChecked():

            if self.lineEdit_nat.text() == None or self.lineEdit_nat.text() == "":
                path = self.lineEdit_nat.text()
                # Open SWAT Scenario folder from GUI
                folder = QtWidgets.QFileDialog.getExistingDirectory(
                    self, "Select SWAT+ 'Scenarios' folder", path
                )

            else:
                # Open SWAT Scenario folder from CLI
                folder = self.lineEdit_nat.text()
                self.pushButton_nat.setEnabled(False)

            # Check if the folder is SWAT+ Scenarios folder
            if folder and folder.endswith("Scenarios"):
                self.lineEdit_nat.setText(folder)
                self.swatplus_nat.setEnabled(True)

                # Clear the comboBox and add folder names
                self.comboBox_scenario_nat.clear()
                subfolders = [f.name for f in os.scandir(folder) if f.is_dir()]
                self.comboBox_scenario_nat.addItems(subfolders)

                # Activate the altered flow inputs
                self.altered.setEnabled(True)

            else:
                # Reset folder, clear comboBox, disable Qframe and show Warning
                folder = None
                self.lineEdit_nat.setText(folder)
                self.comboBox_scenario_nat.clear()
                self.swatplus_nat.setEnabled(False)
                QtWidgets.QMessageBox.warning(
                    self,
                    "Invalid Folder",
                    "Please select the SWAT+ 'Scenarios' folder.",
                )

        if self.radioButton_csv_nat.isChecked():
            self.swatplus_nat.setEnabled(False)
            # Open a CSV file
            fname = QtWidgets.QFileDialog.getOpenFileName(
                self, "Open natural flow CSV file", "", "CSV files (*.csv)"
            )
            self.lineEdit_nat.setText(fname[0])

            # Read the CSV file, check format (columns, daily, gaps) and extract the max and min date
            if fname[0]:
                df = pd.read_csv(fname[0], parse_dates=[0])

                # Check if the CSV has the required columns (also separators)
                if "Date" not in df.columns or "Flow" not in df.columns:
                    QtWidgets.QMessageBox.warning(
                        self,
                        "Invalid CSV Format",
                        "The CSV file must have 'Date' as the first column and 'Flow' as the second column. Please ensure the columns are correctly named and ordered.",
                    )
                    reset = None
                    self.lineEdit_nat.setText(reset)
                    return

                # Check if the 'Date' column is daily
                if not pd.infer_freq(df["Date"]) == "D":
                    QtWidgets.QMessageBox.warning(
                        self,
                        "Invalid Date Frequency",
                        "The 'Date' column must have a daily frequency without any gaps. Please ensure the dates are consecutive and there are no missing days. Also, verify that the last row in the file is correct and that there are no gaps.",
                    )
                    reset = None
                    self.lineEdit_nat.setText(reset)
                    return

                # Check if the 'Flow' column has no gaps or negative values
                if df["Flow"].isnull().any() or (df["Flow"] < 0).any():
                    QtWidgets.QMessageBox.warning(
                        self,
                        "Invalid Flow Data",
                        "The 'Flow' column must have no gaps or negative values. Please ensure there are no missing flow values.",
                    )
                    reset = None
                    self.lineEdit_nat.setText(reset)
                    return

                # Extract the max and min date
                max_date = df.iloc[:, 0].max()
                min_date = df.iloc[:, 0].min()
                if not pd.isnull(max_date):
                    self.DateEdit_finish_year_nat.setDate(QDate(max_date.year, 1, 1))
                if not pd.isnull(min_date):
                    self.DateEdit_start_year_nat.setDate(QDate(min_date.year, 1, 1))

                # Activate the altered flow inputs
                self.altered.setEnabled(True)

    def reset_select_file_nat(self):
        """Pass multiple methods as argument - Reset the natural flow inputs and select the SWAT+ scenario"""
        self.reset_var_nat()
        self.select_file_nat()

    def select_Scenario_nat(self):
        """Select the SWAT+ scenario from the comboBox"""

        # Get the folder and scenario name from GUI
        folder = self.lineEdit_nat.text()
        scenario = self.comboBox_scenario_nat.currentText()

        # Connect to the SQLite database of SWAT+ editor
        sqlite = os.path.join(folder, scenario, "Results", "swatplus_output.sqlite")
        conn = sqlite3.connect(sqlite)
        cursor = conn.cursor()

        # Unique values 'unit' column in 'channel_sd_day' table
        cursor.execute("SELECT DISTINCT unit FROM channel_sd_day")
        channels = [str(row[0]) for row in cursor.fetchall()]  # Convert to string

        # Minimum value 'yr' column in 'channel_sd_day' table
        cursor.execute("SELECT MIN(yr) FROM channel_sd_day")
        min_year = cursor.fetchone()[0]

        # Maximum value 'yr' column in 'channel_sd_day' table
        cursor.execute("SELECT MAX(yr) FROM channel_sd_day")
        max_year = cursor.fetchone()[0]

        # Close the database connection
        conn.close()

        # Add channel 'unit' to comboBox
        self.comboBox_channel_nat.clear()
        self.comboBox_channel_nat.addItems(channels)

        # Add max and min years to QDateEdit
        if min_year:
            self.DateEdit_start_year_nat.setDate(QDate(min_year, 1, 1))
        if max_year:
            self.DateEdit_finish_year_nat.setDate(QDate(max_year, 1, 1))

        # Show a warning if there is no 'channel_sd_day' in scenario
        if channels == []:
            QtWidgets.QMessageBox.warning(
                self,
                "Invalid SWAT+ Scenario",
                f"'channel_sd_day' table not found in {scenario}. Please choose in SWAT+ Editor outputs to print: Daily > Model Components > Channel",
            )

    def select_file_alt(self):
        """Open folder or file for alt input"""

        if self.radioButton_swat_alt.isChecked():
            path = self.lineEdit_nat.text()
            # Open SWAT Scenario folder
            folder = QtWidgets.QFileDialog.getExistingDirectory(
                self, "Select SWAT+ 'Scenarios' folder", path
            )
            # Check if the folder is SWAT+ Scenarios folder
            if folder and folder.endswith("Scenarios"):
                self.lineEdit_alt.setText(folder)
                self.swatplus_alt.setEnabled(True)

                # Clear the comboBox and add folder names
                self.comboBox_scenario_alt.clear()
                subfolders = [f.name for f in os.scandir(folder) if f.is_dir()]
                self.comboBox_scenario_alt.addItems(subfolders)

            else:
                # Reset folder, clear comboBox, disable Qframe and show Warning
                folder = None
                self.lineEdit_alt.setText(folder)
                self.comboBox_scenario_alt.clear()
                self.swatplus_alt.setEnabled(False)
                QtWidgets.QMessageBox.warning(
                    self,
                    "Invalid Folder",
                    "Please select the SWAT+ 'Scenarios' folder.",
                )

        if self.radioButton_csv_alt.isChecked():
            # Deactivate the SWAT+ inputs and reports button
            self.pushButton_reports.setEnabled(False)
            self.swatplus_alt.setEnabled(False)

            # Open a file
            fname = QtWidgets.QFileDialog.getOpenFileName(
                self, "Open altered flow CSV file", "", "CSV files (*.csv)"
            )
            self.lineEdit_alt.setText(fname[0])

            # Read the CSV file, check format (columns, daily, gaps) and extract the max and min date
            if fname[0]:
                df = pd.read_csv(fname[0], parse_dates=[0])

                # Check if the CSV has the required columns
                if "Date" not in df.columns or "Flow" not in df.columns:
                    QtWidgets.QMessageBox.warning(
                        self,
                        "Invalid CSV Format",
                        "The CSV file must have 'Date' as the first column and 'Flow' as the second column. Please ensure the columns are correctly named and ordered.",
                    )
                    reset = None
                    self.lineEdit_alt.setText(reset)
                    return

                # Check if the 'Date' column is daily
                if not pd.infer_freq(df["Date"]) == "D":
                    QtWidgets.QMessageBox.warning(
                        self,
                        "Invalid Date Frequency",
                        "The 'Date' column must have a daily frequency without any gaps. Please ensure the dates are consecutive and there are no missing days.",
                    )
                    reset = None
                    self.lineEdit_alt.setText(reset)
                    return

                # Check if the 'Flow' column has no gaps or negative values
                if df["Flow"].isnull().any() or (df["Flow"] < 0).any():
                    QtWidgets.QMessageBox.warning(
                        self,
                        "Invalid Flow Data",
                        "The 'Flow' column must have no gaps or negative values. Please ensure there are no missing flow values.",
                    )
                    reset = None
                    self.lineEdit_alt.setText(reset)
                    return

                # Extract the max and min date
                max_date = df.iloc[:, 0].max()
                min_date = df.iloc[:, 0].min()
                if not pd.isnull(max_date):
                    self.DateEdit_finish_year_alt.setDate(QDate(max_date.year, 1, 1))
                if not pd.isnull(min_date):
                    self.DateEdit_start_year_alt.setDate(QDate(min_date.year, 1, 1))

                # Activate the reports button
                self.pushButton_reports.setEnabled(True)

    def select_Scenario_alt(self):
        """Select the SWAT+ scenario from the comboBox"""

        # Get the folder and scenario name from GUI
        folder = self.lineEdit_alt.text()
        scenario = self.comboBox_scenario_alt.currentText()

        # Connect to the SQLite database
        sqlite = os.path.join(folder, scenario, "Results", "swatplus_output.sqlite")
        conn = sqlite3.connect(sqlite)
        cursor = conn.cursor()

        # Unique values 'unit' column in 'channel_sd_day' table
        cursor.execute("SELECT DISTINCT unit FROM channel_sd_day")
        channels = [str(row[0]) for row in cursor.fetchall()]  # Convert to string

        # Minimum value 'yr' column in 'channel_sd_day' table
        cursor.execute("SELECT MIN(yr) FROM channel_sd_day")
        min_year = cursor.fetchone()[0]

        # Maximum value 'yr' column in 'channel_sd_day' table
        cursor.execute("SELECT MAX(yr) FROM channel_sd_day")
        max_year = cursor.fetchone()[0]

        # Close the database connection
        conn.close()

        # Add channel 'unit' to comboBox
        self.comboBox_channel_alt.clear()
        self.comboBox_channel_alt.addItems(channels)

        # Add max and min years to comboBox
        if min_year:
            self.DateEdit_start_year_alt.setDate(QDate(min_year, 1, 1))
        if max_year:
            self.DateEdit_finish_year_alt.setDate(QDate(max_year, 1, 1))

        # Show a warning if there is no 'channel_sd_day' in scenario
        if channels == []:
            QtWidgets.QMessageBox.warning(
                self,
                "Invalid SWAT+ Scenario",
                f"'channel_sd_day' table not found in {scenario}. Please choose in SWAT+ Editor outputs to print: Daily > Model Components > Channel",
            )
            self.pushButton_reports.setEnabled(False)
            return

        # Activate the reports button
        self.pushButton_reports.setEnabled(True)

    def generate_reports(self):
        """Generate IAHRIS reports"""

        # Check the start and finish years of both nat and alt period
        start_year_nat = self.DateEdit_start_year_nat.date().year()
        end_year_nat = self.DateEdit_finish_year_nat.date().year()
        start_year_alt = self.DateEdit_start_year_alt.date().year()
        end_year_alt = self.DateEdit_finish_year_alt.date().year()

        if end_year_nat - start_year_nat < 14 or end_year_alt - start_year_alt < 14:
            QtWidgets.QMessageBox.warning(
                self,
                "Invalid selected periods",
                "The selected periods for analysis must cover at least 15 consecutive years. Please adjust the start and end years to ensure that both periods are 15 years or longer.",
            )
            return

        # Check if the 'SWATPlus-IAHRIS' folder and 'temp' folder exist
        if not os.path.exists("C:\\SWATPlus-IAHRIS"):
            QtWidgets.QMessageBox.warning(
                self,
                "Installation error",
                "C:\SWATPlus-IAHRIS not found. Relaunch the installer.",
            )
            return
        # Create temp folder
        if not os.path.exists("C:\\SWATPlus-IAHRIS\\temp"):
            os.makedirs("C:\\SWATPlus-IAHRIS\\temp")

        # Temp folder
        temp_folder = "C:\\SWATPlus-IAHRIS\\temp"

        # Get the directory to save the reports
        report_folder = QtWidgets.QFileDialog.getExistingDirectory(
            self, "Select a folder to save the IAHRIS reports", ""
        ).replace("/", "\\")

        if not report_folder:
            QtWidgets.QMessageBox.warning(
                self,
                "No Folder Selected",
                "Please choose a location to save the IAHRIS reports.",
            )
            return

        self.progressBar.setValue(10)

        # IAHRIS input data for the natural scenario from SWAT+
        if (
            self.radioButton_swat_nat.isChecked()
            and self.comboBox_channel_nat.currentText() != ""
        ):
            # Get the folder and scenario name from GUI
            temp_folder = temp_folder
            folder = self.lineEdit_nat.text()
            scenario_nat = self.comboBox_scenario_nat.currentText()

            # Connect to the SQLite database of SWAT+ editor
            sqlite = os.path.join(
                folder, scenario_nat, "Results", "swatplus_output.sqlite"
            )
            conn = sqlite3.connect(sqlite)

            # Generate a DataFrame with the specified columns filtered by the selected unit
            unit = self.comboBox_channel_nat.currentText()
            start_year = self.DateEdit_start_year_nat.date().year()
            end_year = self.DateEdit_finish_year_nat.date().year()

            query = f"""
            SELECT 
                printf('%02d', day) || '/' || printf('%02d', mon) || '/' || yr AS Date,
                flo_out
            FROM channel_sd_day
            WHERE unit = ? AND yr BETWEEN ? AND ?
            """
            # Read the query into a DataFrame (params connects the '?' in the query with the variables)
            df = pd.read_sql_query(query, conn, params=(unit, start_year, end_year))
            conn.close()

            # Create a header DataFrame
            header = pd.DataFrame(
                [
                    ["DIARIO", "NATURAL", scenario_nat[:12]]
                ],  # IAHRIS limit the scenario name to 12 characters
                columns=["Date", "flo_out", "Scenario_nat"],
            )

            # Concatenate the header df with the query df
            df = pd.concat([header, df], ignore_index=True)

            if df["Date"].iloc[-1] == "31/12/{}".format(end_year):
                # Add one more day to the complete year (to take into account the last year of data)
                next_day = pd.DataFrame(
                    [["01/01/{}".format(end_year + 1), "0.00", ""]], columns=df.columns
                )
                df = pd.concat([df, next_day], ignore_index=True)

            # Save the DataFrame as a CSV file with ';' as the delimiter (required by IAHRIS)
            output_csv_path_nat = os.path.join(temp_folder, f"{scenario_nat}_nat.csv")
            df.to_csv(output_csv_path_nat, sep=";", index=False, header=False)

            self.progressBar.setValue(50)

        # IAHRIS input data for the natural scenario from a CSV file
        if self.radioButton_csv_nat.isChecked():
            temp_folder = temp_folder
            # Read the CSV file
            input_csv_nat = self.lineEdit_nat.text()
            df = pd.read_csv(input_csv_nat)

            # Get the scenario name from the CSV file
            scenario_nat = os.path.splitext(os.path.basename(input_csv_nat))[0]

            # Create a header DataFrame
            header = pd.DataFrame(
                [
                    ["DIARIO", "NATURAL", scenario_nat[:12]]
                ],  # IAHRIS limit the scenario name to 12 characters
                columns=["Date", "Flow", "Scenario_nat"],
            )

            # 'Date' as datetime
            df["Date"] = pd.to_datetime(df["Date"])

            # Get the 'Date' and 'Flow' columns
            df = df[["Date", "Flow"]]

            # Filter df by selected start and finish year
            start_year = self.DateEdit_start_year_nat.date().year()
            end_year = self.DateEdit_finish_year_nat.date().year()
            df = df[
                (df["Date"].dt.year >= start_year) & (df["Date"].dt.year <= end_year)
            ]

            # Format the 'Date' column to DD/MM/YYYY (required by IAHRIS)
            df["Date"] = pd.to_datetime(df["Date"]).dt.strftime("%d/%m/%Y")

            # Concatenate the header df with the csv df (columns 'Date' and 'Flow' should be named as in the header)
            df = pd.concat([header, df], ignore_index=True)

            if df["Date"].iloc[-1] == "31/12/{}".format(end_year):
                # Add one more day to the complete year (to take into account the last year of data)
                next_day = pd.DataFrame(
                    [["01/01/{}".format(end_year + 1), "0.00", ""]], columns=df.columns
                )
                df = pd.concat([df, next_day], ignore_index=True)

            # Save the DataFrame as a CSV file with ';' as the delimiter (required by IAHRIS)
            output_csv_path_nat = os.path.join(temp_folder, f"{scenario_nat}_nat.csv")
            df.to_csv(output_csv_path_nat, sep=";", index=False, header=False)

            self.progressBar.setValue(50)

        # IAHRIS input data for the altered scenario from SWAT+
        if (
            self.radioButton_swat_alt.isChecked()
            and self.comboBox_channel_alt.currentText() != ""
        ):
            # Get the folder and scenario name from GUI
            temp_folder = temp_folder
            folder = self.lineEdit_alt.text()
            scenario_nat = scenario_nat
            scenario_alt = self.comboBox_scenario_alt.currentText()

            # Connect to the SQLite database of SWAT+ editor
            sqlite = os.path.join(
                folder, scenario_alt, "Results", "swatplus_output.sqlite"
            )
            conn = sqlite3.connect(sqlite)

            # Generate a DataFrame with the specified columns filtered by the selected unit
            unit = self.comboBox_channel_alt.currentText()
            start_year = self.DateEdit_start_year_alt.date().year()
            end_year = self.DateEdit_finish_year_alt.date().year()

            query = f"""
            SELECT 
                printf('%02d', day) || '/' || printf('%02d', mon) || '/' || yr AS Date,
                flo_out
            FROM channel_sd_day
            WHERE unit = ? AND yr BETWEEN ? AND ?
            """

            # Read the query into a DataFrame (params connects the '?' in the query with the variables)
            df = pd.read_sql_query(query, conn, params=(unit, start_year, end_year))
            conn.close()

            # Create a header df
            header = pd.DataFrame(
                [
                    ["DIARIO", "ALTERADO", scenario_nat[:12], scenario_alt[:12]]
                ],  # IAHRIS limit the scenario name to 12 characters
                columns=["Date", "flo_out", "Scenario_nat", "Scenario_alt"],
            )

            # Concatenate the header df with the query df
            df = pd.concat([header, df], ignore_index=True)

            if df["Date"].iloc[-1] == "31/12/{}".format(end_year):
                # Add one more day to the complete year (to take into account the last year of data)
                next_day = pd.DataFrame(
                    [["01/01/{}".format(end_year + 1), "0.00", "", ""]],
                    columns=df.columns,
                )
                df = pd.concat([df, next_day], ignore_index=True)

            # Save the DataFrame as a CSV file with ';' as the delimiter (required by IAHRIS)
            output_csv_path_alt = os.path.join(temp_folder, f"{scenario_alt}_alt.csv")
            df.to_csv(output_csv_path_alt, sep=";", index=False, header=False)

            self.progressBar.setValue(80)

        # IAHRIS input data for the altered scenario from a CSV file
        if self.radioButton_csv_alt.isChecked():
            temp_folder = temp_folder
            # Read the CSV file
            input_csv_alt = self.lineEdit_alt.text()
            df = pd.read_csv(input_csv_alt)

            # Get the scenario name from the CSV file
            scenario_nat = scenario_nat
            scenario_alt = os.path.splitext(os.path.basename(input_csv_alt))[0]

            # Create a header df
            header = pd.DataFrame(
                [
                    ["DIARIO", "ALTERADO", scenario_nat[:12], scenario_alt[:12]]
                ],  # IAHRIS limit the scenario name to 12 characters
                columns=["Date", "Flow", "Scenario_nat", "Scenario_alt"],
            )

            # 'Date' as datetime
            df["Date"] = pd.to_datetime(df["Date"])

            # Get the 'Date' and 'Flow' columns
            df = df[["Date", "Flow"]]

            # Filter df by selected start and finish year
            start_year = self.DateEdit_start_year_alt.date().year()
            end_year = self.DateEdit_finish_year_alt.date().year()
            df = df[
                (df["Date"].dt.year >= start_year) & (df["Date"].dt.year <= end_year)
            ]

            # Format the 'Date' column to DD/MM/YYYY (required by IAHRIS)
            df["Date"] = pd.to_datetime(df["Date"]).dt.strftime("%d/%m/%Y")

            # Concatenate the header df with the csv df (columns 'Date' and 'Flow' should be named as in the header)
            df = pd.concat([header, df], ignore_index=True)

            if df["Date"].iloc[-1] == "31/12/{}".format(end_year):
                # Add one more day to the complete year (to take into account the last year of data)
                next_day = pd.DataFrame(
                    [["01/01/{}".format(end_year + 1), "0.00", "", ""]],
                    columns=df.columns,
                )
                df = pd.concat([df, next_day], ignore_index=True)

            # Save the DataFrame as a CSV file with ';' as the delimiter (required by IAHRIS)
            output_csv_path_alt = os.path.join(temp_folder, f"{scenario_alt}_alt.csv")
            df.to_csv(output_csv_path_alt, sep=";", index=False, header=False)

            self.progressBar.setValue(80)

        current_date = datetime.now().strftime("%Y-%m-%d_%H-%M")
        project_name = current_date

        # Generate the .bat file of IAHRIS
        scenario_nat_short = scenario_nat[
            :12
        ]  # IAHRIS limit the scenario name to 12 characters
        scenario_alt_short = scenario_alt[
            :12
        ]  # IAHRIS limit the scenario name to 12 characters
        bat_content = f"""cd C:\\SWATPlus-IAHRIS\\IAHRIS4.0
chcp 65001
:: Lineas de Carga de Datos.
IAHRIS.exe CD /t:P /np:"{scenario_nat_short}" /d:"Punto Carga Masiva" /p:"{project_name}" /dp:"SWATPlus-IAHRIS" /rn:"Natural" /ra:"nat" /fe:"{output_csv_path_nat}" /mi:01
IAHRIS.exe CD /t:A /np:"{scenario_nat_short}" /na:"{scenario_alt_short}" /d:"Alt Carga Masiva" /p:"{project_name}" /dp:"SWATPlus-IAHRIS" /rn:"Alterado" /ra:"alt" /fe:"{output_csv_path_alt}"

:: Lineas para Generar informe de Salida.
IAHRIS.exe GIS /t:A /np:"{scenario_nat_short}" /na:"{scenario_alt_short}" /p:"{project_name}" /fs:"{report_folder}" -cvh -cas
"""
        bat_file_path = os.path.join(temp_folder, "generate_report.bat")
        with open(bat_file_path, "w", encoding="utf-8") as bat_file:
            bat_file.write(bat_content)

        # Launch the .bat file
        subprocess.run(
            [bat_file_path], shell=True, creationflags=subprocess.CREATE_NO_WINDOW
        )
        self.progressBar.setValue(100)

        # Remove the temp folder
        shutil.rmtree("C:\\SWATPlus-IAHRIS\\temp")

        self.progressBar.setValue(90)

        # Get the list of .xlsx files in the report_folder
        xlsx_files = glob.glob(os.path.join(report_folder, "*.xlsx"))

        try:
            if xlsx_files:
                # Get the last generated .xlsx file
                last_generated_xlsx = max(xlsx_files, key=os.path.getctime)

                # Open the workbook and select the first sheet
                app = xlwings.App(visible=False)
                workbook = app.books.open(last_generated_xlsx)
                sheet = workbook.sheets[0]

                # Change the values of the cells
                sheet["AA1"].value = "Nat_F"
                sheet["AA2"].value = "Alt_F"
                sheet["AB1"].value = "Natural Flow"
                sheet["AB2"].value = "Altered Flow"
                sheet["E3"].value = ""
                sheet["E4"].value = ""

                # Save the changes to the workbook
                workbook.save()
                workbook.close()
                app.quit()
        except:
            QtWidgets.QMessageBox.warning(
                self,
                "Excel File Access Error",
                "Please ensure that all Excel files are closed before continuing. If the problem persists, check for any background Excel processes and try again.",
            )
            self.progressBar.setValue(0)
            return

        # # Open the report folder
        # os.startfile(report_folder)

        self.progressBar.setValue(0)

        self.reports_window = ReportsWindow(last_generated_xlsx, report_folder)
        self.reports_window.show()


class ReportsWindow(QtWidgets.QMainWindow):
    def __init__(self, last_generated_xlsx, report_folder):
        """Constructor of the SWATPlus-IAHRIS software - Report Window"""
        super().__init__()
        uic.loadUi(resource_path("reports.ui"), self)

        # Variables from the main window
        self.last_generated_xlsx = last_generated_xlsx  # Master report file
        self.report_folder = report_folder  # Selected folder to save the reports

        self.pushButton_print.clicked.connect(self.on_print_button_clicked)

    def extract_selected_sheets_to_excel(self, selected_sheet_names, output_excel_path):
        """Extract selected reports from the master report."""

        app = xlwings.App(visible=False)
        try:
            wb = app.books.open(self.last_generated_xlsx)
            # print([s.name for s in wb.sheets])
            # Create a new workbook
            new_wb = app.books.add()

            # Copy selected sheets
            for sheet_name in selected_sheet_names:
                if sheet_name in [s.name for s in wb.sheets]:
                    wb.sheets[sheet_name].copy(
                        after=new_wb.sheets[-1] if new_wb.sheets else None
                    )
            # Delete default sheet
            new_wb.sheets[0].delete()

            # Save the new workbook
            new_wb.save(output_excel_path)
            new_wb.close()
            wb.close()
        finally:
            app.quit()

    def rename_sheets_in_excel(self, output_excel_path):

        rename_dict = {
            "REPORTS": "Reports",
            "Informe nº1": "Report_n1",
            "Informe nº1a": "Report_n1a",
            "Informe nº 1b": "Report_n1b",
            "Informe nº 2": "Report_n2",
            "Informe nº2a": "Report_n2a",
            "Informe nº3": "Report_n3",
            "Informe nº3a": "Report_n3a",
            "Informe nº 3b": "Report_n3b",
            "Informe nº3c": "Report_n3c",
            "Informe nº4": "Report_n4",
            "Informe nº5": "Report_n5",
            "Informe nº 6": "Report_n6",
            "Informe nº 6a": "Report_n6a",
            "Informe nº6b": "Report_n6b",
            "Informe nº6c": "Report_n6c",
            "Informe nº6d": "Report_n6d",
            "Informe nº6e": "Report_n6e",
            "Informe nº 7a": "Report_n7a",
            "Informe nº 7c": "Report_n7c",
            "Informe nº 7d": "Report_n7d",
            "Informe nº 8": "Report_n8",
            "Informe nº8b": "Report_n8b",
            "Informe nº10a": "Report_n10a",
            "Informe nº 10c": "Report_n10c",
        }

        app = xlwings.App(visible=False)
        try:
            wb = app.books.open(output_excel_path)
            for old_name, new_name in rename_dict.items():
                for sheet in wb.sheets:
                    if sheet.name == old_name:
                        sheet.name = new_name
            wb.save()
            wb.close()
        finally:
            app.quit()

    def on_print_button_clicked(self):
        """Extract reports based on the selected checkboxes (themes)."""
        if self.checkBox_nat.isChecked():
            selected_sheet_names = [
                "Informe nº1",
                "Informe nº 2",
                "Informe nº2a",
                "Informe nº4",
            ]
            output_excel_path = (
                self.report_folder
                + "\\SWATPlus-IAHRIS_Natural_Flow_Characterization.xlsx"
            )
            self.extract_selected_sheets_to_excel(
                selected_sheet_names, output_excel_path
            )
            self.rename_sheets_in_excel(output_excel_path)
            self.label_nat.setText("✓")
        self.progressBar.setValue(10)

        if self.checkBox_alt.isChecked():
            selected_sheet_names = [
                "Informe nº1a",
                "Informe nº3",
                "Informe nº3a",
                "Informe nº5",
            ]
            output_excel_path = (
                self.report_folder
                + "\\SWATPlus-IAHRIS_Altered_Flow_Characterization.xlsx"
            )
            self.extract_selected_sheets_to_excel(
                selected_sheet_names, output_excel_path
            )
            self.rename_sheets_in_excel(output_excel_path)
            self.label_alt.setText("✓")
        self.progressBar.setValue(30)
        if self.checkBox_nat_alt.isChecked():
            selected_sheet_names = [
                "Informe nº 1b",
                "Informe nº 3b",
                "Informe nº3c",
                "Informe nº 8",
                "Informe nº8a",
            ]
            output_excel_path = (
                self.report_folder
                + "\\SWATPlus-IAHRIS_Natural-Altered_Flow_Comparison.xlsx"
            )
            self.extract_selected_sheets_to_excel(
                selected_sheet_names, output_excel_path
            )
            self.rename_sheets_in_excel(output_excel_path)
            self.label_nat_alt.setText("✓")
        self.progressBar.setValue(40)
        if self.checkBox_curves.isChecked():
            selected_sheet_names = [
                "Informe nº 6",
                "Informe nº 6a",
                "Informe nº6b",
                "Informe nº6c",
                "Informe nº6d",
                "Informe nº6e",
            ]
            output_excel_path = (
                self.report_folder + "\\SWATPlus-IAHRIS_Flow_Rates_Duration_Curves.xlsx"
            )
            self.extract_selected_sheets_to_excel(
                selected_sheet_names, output_excel_path
            )
            self.rename_sheets_in_excel(output_excel_path)
            self.label_curves.setText("✓")
        self.progressBar.setValue(60)
        if self.checkBox_habitual.isChecked():
            selected_sheet_names = [
                "Informe nº 7a",
                "Informe nº 7c",
            ]
            output_excel_path = (
                self.report_folder + "\\SWATPlus-IAHRIS_IHA_Habitual_Values.xlsx"
            )
            self.extract_selected_sheets_to_excel(
                selected_sheet_names, output_excel_path
            )
            self.rename_sheets_in_excel(output_excel_path)
            self.label_habitual.setText("✓")
        self.progressBar.setValue(70)
        if self.checkBox_floods.isChecked():
            selected_sheet_names = [
                "Informe nº 7d",
            ]
            output_excel_path = (
                self.report_folder + "\\SWATPlus-IAHRIS_IHA_Floods_Droughts.xlsx"
            )
            self.extract_selected_sheets_to_excel(
                selected_sheet_names, output_excel_path
            )
            self.rename_sheets_in_excel(output_excel_path)
            self.label_floods.setText("✓")
        self.progressBar.setValue(90)
        if self.checkBox_sign.isChecked():
            selected_sheet_names = ["Informe nº10a", "Informe nº 10c"]
            output_excel_path = (
                self.report_folder
                + "\\SWATPlus-IAHRIS_IHA_Environmental_Significance.xlsx"
            )
            self.extract_selected_sheets_to_excel(
                selected_sheet_names, output_excel_path
            )
            self.rename_sheets_in_excel(output_excel_path)
            self.label_sign.setText("✓")
        self.progressBar.setValue(100)

        # Open the report folder
        os.startfile(self.report_folder)

        self.progressBar.setValue(0)


if __name__ == "__main__":
    import sys

    # Create the application
    app = QtWidgets.QApplication(sys.argv)

    # Create an instance of the MainWindow class
    window = MainWindow()

    # Check if a folder path is passed as a command-line argument
    if len(sys.argv) > 1:
        folder = sys.argv[1]  # Get the folder path from the command-line argument
        if os.path.exists(folder) and os.path.isdir(folder):
            window.lineEdit_nat.setText(folder)  # Set the folder path in the GUI
            window.select_file_nat()
        else:
            pass

    # Show the main window
    window.show()

    # Run the application's event loop
    sys.exit(app.exec())
