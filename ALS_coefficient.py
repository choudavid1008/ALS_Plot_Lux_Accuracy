import sys
import os
import shutil
import pandas as pd
import numpy as np
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QTabWidget, QTableWidget, QTableWidgetItem,
    QLineEdit, QLabel, QFileDialog, QFormLayout, QGroupBox,
    QScrollArea, QMessageBox, QTextEdit
)
from PySide6.QtCore import Qt

# Try to import official project modules if available on user's machine
try:
    import common.columns as columns
    import processor.raw_data_processor as raw_data_processor
    import processor.reduction_processor as reduction_processor
    import processor.baseline_processor as baseline_processor
    import processor.normalize_processor as normalize_processor
    import processor.regression_processor as regression_processor
    import util.log_util as log_util
    from common.models import RawDataFile
    OFFICIAL_PIPELINE_AVAILABLE = True
except ImportError:
    OFFICIAL_PIPELINE_AVAILABLE = False

class ALS_CoefficientApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ALS Coefficient Calibration Tool")
        self.resize(1200, 800)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)

        # Top Buttons Layout
        self.top_layout = QHBoxLayout()
        self.load_btn = QPushButton("Load File(s) (讀檔)")
        self.load_btn.clicked.connect(self.load_files)
        self.top_layout.addWidget(self.load_btn)

        self.run_btn = QPushButton("Run (執行)")
        self.run_btn.clicked.connect(self.run_calibration)
        self.top_layout.addWidget(self.run_btn)
        self.main_layout.addLayout(self.top_layout)

        # Tab Widget
        self.tabs = QTabWidget()
        self.main_layout.addWidget(self.tabs)

        # Tab 1: Config & Data View
        self.data_tab = QWidget()
        self.data_layout = QHBoxLayout(self.data_tab)
        self.tabs.addTab(self.data_tab, "Data & Configuration (資料與設定)")

        # Left Column: Configuration
        self.config_group = QGroupBox("Calibration Configuration")
        self.config_layout = QFormLayout(self.config_group)
        self.config_group.setFixedWidth(350)

        self.gain_input = QLineEdit("256")
        self.sample_time_input = QLineEdit("500")
        self.samples_input = QLineEdit("300")

        self.config_layout.addRow("GAIN:", self.gain_input)
        self.config_layout.addRow("SAMPLE_TIME:", self.sample_time_input)
        self.config_layout.addRow("SAMPLES:", self.samples_input)

        # Target Displays
        self.detected_ccts_label = QLabel("None")
        self.detected_ccts_label.setWordWrap(True)
        self.config_layout.addRow("Detected CCTs:", self.detected_ccts_label)

        self.detected_lux_label = QTextEdit()
        self.detected_lux_label.setReadOnly(True)
        self.detected_lux_label.setFixedHeight(150)
        self.config_layout.addRow("Detected LUX Map:", self.detected_lux_label)

        # Mode Indicator
        self.mode_label = QLabel()
        if OFFICIAL_PIPELINE_AVAILABLE:
            self.mode_label.setText("Pipeline: <font color='green'><b>Official Modules</b></font>")
        else:
            self.mode_label.setText("Pipeline: <font color='orange'><b>Fallback Adaptive Engine</b></font>")
        self.config_layout.addRow("Mode:", self.mode_label)

        self.data_layout.addWidget(self.config_group)

        # Right Column: Loaded Data Table View
        self.table_widget = QTableWidget()
        self.data_layout.addWidget(self.table_widget)

        # Tab 2: Execution Output View
        self.output_tab = QWidget()
        self.output_layout = QVBoxLayout(self.output_tab)
        self.tabs.addTab(self.output_tab, "Execution Console (執行輸出)")

        self.console_output = QTextEdit()
        self.console_output.setReadOnly(True)
        self.output_layout.addWidget(self.console_output)

        # State Variables
        self.loaded_files = []
        self.combined_df = None
        self.target_ccts = []
        self.target_lux_map = {}

    def load_files(self):
        file_paths, _ = QFileDialog.getOpenFileNames(
            self, "Open Raw Data Files", "", "Excel/CSV Files (*.xlsx *.xls *.csv)"
        )
        if not file_paths:
            return

        self.loaded_files = file_paths
        dfs = []

        for path in file_paths:
            try:
                if path.endswith(".csv"):
                    df = pd.read_csv(path)
                else:
                    df = pd.read_excel(path, header=0)
                dfs.append(df)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Could not load file {os.path.basename(path)}: {e}")
                return

        if dfs:
            self.combined_df = pd.concat(dfs, ignore_index=True)
            self.populate_table(self.combined_df)
            self.auto_calculate_targets(self.combined_df)
            QMessageBox.information(self, "Success", f"Successfully loaded {len(file_paths)} files.")

    def populate_table(self, df):
        preview_limit = min(len(df), 100)
        self.table_widget.setRowCount(preview_limit)
        self.table_widget.setColumnCount(len(df.columns))
        self.table_widget.setHorizontalHeaderLabels(list(df.columns))

        for i in range(preview_limit):
            for j in range(len(df.columns)):
                val = str(df.iloc[i, j]) if pd.notna(df.iloc[i, j]) else ""
                self.table_widget.setItem(i, j, QTableWidgetItem(val))

    def auto_calculate_targets(self, df):
        # Standardize column casing
        df.columns = [c.strip().upper() for c in df.columns]

        cct_col = None
        lux_col = None

        for col in df.columns:
            if "CCT" in col:
                cct_col = col
            if "LUX" in col:
                lux_col = col

        if not cct_col or not lux_col:
            self.detected_ccts_label.setText("CCT or LUX columns not found!")
            self.detected_lux_label.setPlainText("")
            return

        # Extract unique CCTs
        self.target_ccts = sorted(df[cct_col].dropna().unique().tolist())
        self.detected_ccts_label.setText(", ".join(map(str, self.target_ccts)))

        # Build TARGET_LUX_MAP
        self.target_lux_map = {}
        lux_map_text = ""
        for cct in self.target_ccts:
            lux_vals = sorted(df[df[cct_col] == cct][lux_col].dropna().unique().tolist())
            self.target_lux_map[cct] = lux_vals
            lux_map_text += f"{cct}K: {lux_vals}\n"

        self.detected_lux_label.setPlainText(lux_map_text)

    def run_calibration(self):
        if not self.loaded_files or self.combined_df is None:
            QMessageBox.warning(self, "Warning", "Please load data files first.")
            return

        try:
            gain_val = float(self.gain_input.text())
            sample_time_val = float(self.sample_time_input.text())
            samples_val = float(self.samples_input.text())
        except ValueError:
            QMessageBox.critical(self, "Error", "GAIN, SAMPLE_TIME, and SAMPLES must be numeric values.")
            return

        self.console_output.clear()
        self.console_output.append("***** Running calibration pipeline... *****")

        # -----------------------------------------------------
        # OPTION 1: Execute using the codebase's official modules
        # -----------------------------------------------------
        if OFFICIAL_PIPELINE_AVAILABLE:
            self.console_output.append("Running with OFFICIAL Codebase Processor modules...")
            try:
                # Set dynamic parameters to official processor globals
                normalize_processor.GAIN = gain_val
                normalize_processor.TINT = ((sample_time_val + 1) * (samples_val + 1)) / 720
                reduction_processor.TARGET_CCTS = self.target_ccts
                reduction_processor.TARGET_LUX_MAP = self.target_lux_map

                raw_files_list = []
                for path in self.loaded_files:
                    raw_files_list.append(RawDataFile(file_name=path, header_row=0))
                raw_data_processor.RAW_DATA_FILES = raw_files_list

                scaled_dfs = pd.DataFrame()
                for raw_file in raw_data_processor.RAW_DATA_FILES:
                    raw_data_df = raw_data_processor.extract_raw_data_df(raw_file)
                    if raw_data_df is None:
                        continue
                    has_exposure_info = raw_data_processor.has_exposure_info(raw_data_df)

                    # Call standard codebase feature-building process
                    all_machine_dfs = []
                    if has_exposure_info:
                        RAW_DATA_DF_COLS = columns.RAW_DATA_DF_COLS_EXPOSURE
                    else:
                        RAW_DATA_DF_COLS = columns.RAW_DATA_DF_COLS

                    for machine_id, machine_cols in RAW_DATA_DF_COLS.items():
                        single_machine_df = raw_data_processor.extract_single_machine_df(has_exposure_info, raw_data_df, machine_cols)
                        single_machine_df = reduction_processor.data_reduction_measurement(has_exposure_info, machine_id, single_machine_df)
                        single_machine_df = reduction_processor.build_calibration_features(machine_id, single_machine_df)
                        if not single_machine_df.empty:
                            all_machine_dfs.append(single_machine_df)

                    reduction_df = reduction_processor.build_reduction_df(raw_file.file_name, all_machine_dfs, has_exposure_info)
                    scaled_df = normalize_processor.scale_reduction_df(has_exposure_info, reduction_df)

                    if scaled_dfs.empty:
                        scaled_dfs = scaled_df.copy()
                    else:
                        scaled_dfs = pd.concat([scaled_dfs, scaled_df], ignore_index=True)

                prediction_df = regression_processor.initialize_prediction_df(scaled_dfs)
                label, prediction_df, coefficients = regression_processor.process_regression_pipeline(scaled_dfs, prediction_df)

                # Output calculated coefficients to the console
                self.console_output.append("\n========================================")
                self.console_output.append(f"Official Target: {label}")
                self.console_output.append("Calculated Coefficients:")
                self.console_output.append(f"Cr: {coefficients[0]:.4f}")
                self.console_output.append(f"Cg: {coefficients[1]:.4f}")
                self.console_output.append(f"Cb: {coefficients[2]:.4f}")
                self.console_output.append(f"Cc: {coefficients[3]:.4f}")
                self.console_output.append(f"Cwb: {coefficients[4]:.4f}")
                self.console_output.append("========================================\n")

            except Exception as e:
                self.console_output.append(f"Official execution encountered an error: {e}. Falling back to Adaptive Engine...")
                self.run_fallback_engine(gain_val, sample_time_val, samples_val)

        # -----------------------------------------------------
        # OPTION 2: Fallback adaptive engine (For non-module environments)
        # -----------------------------------------------------
        else:
            self.console_output.append("Running with Fallback Adaptive Engine...")
            self.run_fallback_engine(gain_val, sample_time_val, samples_val)

        # Copy loaded files to the output_file directory
        os.makedirs("output_file", exist_ok=True)
        for path in self.loaded_files:
            dest = os.path.join("output_file", os.path.basename(path))
            try:
                shutil.copy(path, dest)
                self.console_output.append(f"Copied source file to: {dest}")
            except Exception as e:
                self.console_output.append(f"Could not copy {os.path.basename(path)}: {e}")

        self.console_output.append("\nProcess complete!")
        self.tabs.setCurrentIndex(1)

    def run_fallback_engine(self, gain, sample_time, samples):
        df = self.combined_df.copy()
        df.columns = [c.strip().upper() for c in df.columns]
        required_cols = ["CCT", "LUX", "RED", "GREEN", "BLUE", "CLEAR", "WB"]
        missing_cols = [c for c in required_cols if c not in df.columns]

        if missing_cols:
            self.console_output.append(f"Error: Missing required columns: {', '.join(missing_cols)}")
            return

        tint = ((sample_time + 1) * (samples + 1)) / 720
        self.console_output.append(f"Calculated TINT: {tint:.4f}")

        filtered_rows = []
        for index, row in df.iterrows():
            cct = row["CCT"]
            lux = row["LUX"]
            if cct in self.target_lux_map:
                if any(np.isclose(lux, target_lux) for target_lux in self.target_lux_map[cct]):
                    filtered_rows.append(row)

        if not filtered_rows:
            self.console_output.append("No matching records found based on CCT and LUX mapping.")
            return

        f_df = pd.DataFrame(filtered_rows)
        self.console_output.append(f"Processing {len(f_df)} matching data rows.")

        scale = (tint * gain) / 256.0
        self.console_output.append(f"Normalisation scale factor: {scale:.4f}")

        for ch in ["RED", "GREEN", "BLUE", "CLEAR", "WB"]:
            f_df[ch] = f_df[ch] / scale

        Y = f_df["LUX"].values
        X = f_df[["RED", "GREEN", "BLUE", "CLEAR", "WB"]].values
        weights = np.where(Y > 0, 1.0 / Y, 1.0)
        W = np.diag(weights)

        try:
            XTWX = X.T @ W @ X
            XTWY = X.T @ W @ Y
            beta = np.linalg.solve(XTWX, XTWY)
        except np.linalg.LinAlgError:
            beta, _, _, _ = np.linalg.lstsq(X, Y, rcond=None)

        self.console_output.append("\n========================================")
        self.console_output.append("Calculated Coefficients:")
        self.console_output.append(f"Cr: {beta[0]:.4f}")
        self.console_output.append(f"Cg: {beta[1]:.4f}")
        self.console_output.append(f"Cb: {beta[2]:.4f}")
        self.console_output.append(f"Cc: {beta[3]:.4f}")
        self.console_output.append(f"Cwb: {beta[4]:.4f}")
        self.console_output.append("========================================\n")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ALS_CoefficientApp()
    window.show()
    sys.exit(app.exec())
