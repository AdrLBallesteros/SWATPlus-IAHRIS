from PyQt6 import QtWidgets, uic
from PyQt6.QtCore import QDate
import os
import sqlite3
import pandas as pd
import subprocess
import shutil
from datetime import datetime


# Define a MainWindow class that inherits from QMainWindow
class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        """Constructor of the SWATPlus-IAHRIS software"""

        super().__init__()

        # Load the UI file
        uic.loadUi("GUI.ui", self)

        # Connect click events to a function
        self.pushButton_nat.clicked.connect(self.select_file_nat)
        self.comboBox_scenario_nat.activated.connect(self.select_Scenario_nat)

        self.pushButton_alt.clicked.connect(self.select_file_alt)
        self.comboBox_scenario_alt.activated.connect(self.select_Scenario_alt)

        self.pushButton_reports.clicked.connect(self.generate_reports)

        # Deactivate SWAT+ parts of the GUI
        self.swatplus_nat.setEnabled(False)
        self.swatplus_alt.setEnabled(False)
        # Deactivate altered flow inputs
        self.altered.setEnabled(False)
        # Deactivate the reports button
        self.pushButton_reports.setEnabled(False)

    def select_file_nat(self):
        """Open folder or file for nat input"""

        if self.radioButton_swat_nat.isChecked():
            # Open SWAT Scenario folder
            folder = QtWidgets.QFileDialog.getExistingDirectory(
                self, "Select SWAT+ 'Scenarios' folder", ""
            )
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
                        "The 'Date' column must have a daily frequency without any gaps. Please ensure the dates are consecutive and there are no missing days.",
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
            # Open SWAT Scenario folder
            folder = QtWidgets.QFileDialog.getExistingDirectory(
                self, "Select SWAT+ 'Scenarios' folder", ""
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
                    [["01/01/{}".format(end_year + 1), "", ""]], columns=df.columns
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
                    [["01/01/{}".format(end_year + 1), "", ""]], columns=df.columns
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
                    [["01/01/{}".format(end_year + 1), "", "", ""]], columns=df.columns
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
                    [["01/01/{}".format(end_year + 1), "", "", ""]], columns=df.columns
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

        # Open the report folder
        os.startfile(report_folder)

        self.progressBar.setValue(0)


if __name__ == "__main__":
    import sys

    # Create the application
    app = QtWidgets.QApplication(sys.argv)

    # Create an instance of the MainWindow class
    window = MainWindow()

    # Show the main window
    window.show()

    # Run the application's event loop
    sys.exit(app.exec())
