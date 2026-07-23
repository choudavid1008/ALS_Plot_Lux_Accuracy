import sys
import os
import shutil
import pandas as pd
import numpy as np
import ast
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
        self.resize(1200, 850)

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

        # Left Column: Scrollable Configuration Container
        self.scroll_config = QScrollArea()
        self.scroll_config.setWidgetResizable(True)
        self.scroll_config.setFixedWidth(380)
        self.config_widget = QWidget()
        self.config_vbox = QVBoxLayout(self.config_widget)
        self.scroll_config.setWidget(self.config_widget)

        # Form Layout inside Config Group
        self.config_group = QGroupBox("Calibration Settings")
        self.config_layout = QFormLayout(self.config_group)

        self.gain_input = QLineEdit("256")
        self.sample_time_input = QLineEdit("500")
        self.samples_input = QLineEdit("300")

        self.config_layout.addRow("GAIN:", self.gain_input)
        self.config_layout.addRow("SAMPLE_TIME:", self.sample_time_input)
        self.config_layout.addRow("SAMPLES:", self.samples_input)

        # Target Displays (Now Clean Comma-Separated CCT QLineEdit)
        self.detected_ccts_input = QLineEdit()
        self.config_layout.addRow("Detected CCTs (可編輯):", self.detected_ccts_input)
        self.config_vbox.addWidget(self.config_group)

        # Dedicated, dynamically-updating GroupBox for LUX rows
        self.lux_map_group = QGroupBox("Detected LUX Map Settings")
        self.lux_map_layout = QFormLayout(self.lux_map_group)
        self.config_vbox.addWidget(self.lux_map_group)

        # Mode Indicator & Stretch
        self.mode_label = QLabel()
        if OFFICIAL_PIPELINE_AVAILABLE:
            self.mode_label.setText("Pipeline: <font color='green'><b>Official Modules</b></font>")
        else:
            self.mode_label.setText("Pipeline: <font color='orange'><b>Fallback Adaptive Engine</b></font>")
        self.config_vbox.addWidget(self.mode_label)
        self.config_vbox.addStretch()

        self.data_layout.addWidget(self.scroll_config)

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
        self.lux_input_fields = {} # dict mapping index -> QLineEdit field

        # Keep track of detected columns
        self.cct_col_name = None
        self.lux_col_name = None
        self.red_col_name = None
        self.green_col_name = None
        self.blue_col_name = None
        self.clear_col_name = None
        self.wb_col_name = None

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
        # Strict exact-match columns scan (Case-insensitive, stripped of spaces)
        original_cols = list(df.columns)
        norm_cols = [c.strip().upper() for c in original_cols]

        self.cct_col_name = None
        self.lux_col_name = None
        self.red_col_name = None
        self.green_col_name = None
        self.blue_col_name = None
        self.clear_col_name = None
        self.wb_col_name = None

        for idx, col in enumerate(norm_cols):
            if col == "CCT":
                self.cct_col_name = original_cols[idx]
            if col == "LUX":
                self.lux_col_name = original_cols[idx]
            if col == "RED":
                self.red_col_name = original_cols[idx]
            if col == "GREEN":
                self.green_col_name = original_cols[idx]
            if col == "BLUE":
                self.blue_col_name = original_cols[idx]
            if col == "CLEAR":
                self.clear_col_name = original_cols[idx]
            if col == "WB":
                self.wb_col_name = original_cols[idx]

        if not self.cct_col_name or not self.lux_col_name:
            self.detected_ccts_input.setText("CCT or LUX columns not found!")
            # Clear layout
            while self.lux_map_layout.count() > 0:
                child = self.lux_map_layout.takeAt(0)
                if child.widget():
                    child.widget().deleteLater()
            return

        # Extract unique CCTs (comma-separated display, e.g. 2300, 2800, 6500)
        self.target_ccts = sorted([int(x) for x in df[self.cct_col_name].dropna().unique() if pd.notna(x)])
        self.detected_ccts_input.setText(", ".join(map(str, self.target_ccts)))

        # Clear existing dynamic LUX Map Rows
        while self.lux_map_layout.count() > 0:
            child = self.lux_map_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

        self.lux_input_fields = {}

        # Rebuild dedicated rows for each CCT (labeled TARGET_CCTS[idx] / e.g. TARGET_CCTS[0] -> 2300)
        for idx, cct in enumerate(self.target_ccts):
            lux_vals = sorted([float(x) for x in df[(df[self.cct_col_name] == cct) & (df[self.lux_col_name].notna())][self.lux_col_name].unique()])

            # Format nicely as a clean comma-separated list of numbers
            lux_str = ", ".join(map(str, lux_vals))

            lux_field = QLineEdit(lux_str)
            label_text = f"TARGET_CCTS[{idx}] ({cct}K):"
            self.lux_map_layout.addRow(label_text, lux_field)

            # Store references
            self.lux_input_fields[idx] = (cct, lux_field)

    def run_calibration(self):
        if not self.loaded_files or self.combined_df is None:
            QMessageBox.warning(self, "Warning", "Please load data files first.")
            return

        # Dynamically retrieve custom user edits from each CCT and row LUX input field
        try:
            # Parse CCTs
            cct_text = self.detected_ccts_input.text()
            self.target_ccts = sorted([int(x.strip()) for x in cct_text.split(",") if x.strip()])

            # Parse target LUX mappings for each dynamic row
            self.target_lux_map = {}
            for idx, (original_cct, lux_field) in self.lux_input_fields.items():
                if idx < len(self.target_ccts):
                    cct_key = self.target_ccts[idx]
                    lux_text = lux_field.text()
                    lux_vals = sorted([float(x.strip()) for x in lux_text.split(",") if x.strip()])
                    self.target_lux_map[cct_key] = lux_vals
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to parse target variables from inputs: {e}\nCCT and Lux values should be numbers separated by commas.")
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

        # Ensure dynamic system folders exist
        os.makedirs("input_file", exist_ok=True)
        os.makedirs("output_file", exist_ok=True)

        # Copy selected files to input_file directory so codebase modules can resolve them smoothly
        for path in self.loaded_files:
            dest_input = os.path.join("input_file", os.path.basename(path))
            try:
                shutil.copy(path, dest_input)
            except Exception as e:
                print(f"Failed to copy to input_file: {e}")

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

                reduction_processor.OLD_REDUCTION_DATA_IS_APPEND = False
                reduction_processor.OLD_REDUCTION_DATA_IS_DEDUPLICATE = False
                reduction_processor.OLD_REDUCTION_DATA_IS_SORT = True

                baseline_processor.BASELINE_CCT = 0
                baseline_processor.BASELINE_LUX = 0

                regression_processor.PREDICTION_DATA_FILE = "PredictionData.xlsx"

                raw_files_list = []
                for path in self.loaded_files:
                    # Provide base filename first as expected by official project structure
                    filename = os.path.basename(path)
                    raw_files_list.append(RawDataFile(file_name=filename, header_row=0))
                raw_data_processor.RAW_DATA_FILES = raw_files_list

                scaled_dfs = pd.DataFrame()
                reduction_dfs = pd.DataFrame()
                for raw_file in raw_data_processor.RAW_DATA_FILES:
                    raw_data_df = None
                    try:
                        raw_data_df = raw_data_processor.extract_raw_data_df(raw_file)
                    except Exception:
                        # Try with absolute path if base filename failed
                        try:
                            raw_file.file_name = os.path.abspath(os.path.join("input_file", raw_file.file_name))
                            raw_data_df = raw_data_processor.extract_raw_data_df(raw_file)
                        except Exception:
                            pass

                    if raw_data_df is None:
                        continue

                    has_exposure_info = raw_data_processor.has_exposure_info(raw_data_df)

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

                    if reduction_dfs.empty:
                        reduction_dfs = reduction_df.copy()
                    else:
                        reduction_dfs = pd.concat([reduction_dfs, reduction_df], ignore_index=True)

                    scaled_df = normalize_processor.scale_reduction_df(has_exposure_info, reduction_df)

                    if scaled_dfs.empty:
                        scaled_dfs = scaled_df.copy()
                    else:
                        scaled_dfs = pd.concat([scaled_dfs, scaled_df], ignore_index=True)

                reduction_processor.output_final_reduction_file(reduction_dfs)

                prediction_df = regression_processor.initialize_prediction_df(scaled_dfs)
                label, prediction_df, coefficients = regression_processor.process_regression_pipeline(scaled_dfs, prediction_df)

                regression_processor.output_prediction_file(prediction_df)

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

        # Enforce exact column name resolution (Case-insensitive exact matches)
        original_cols = list(df.columns)
        norm_cols = [c.strip().upper() for c in original_cols]

        def find_exact_col(sub):
            if sub in norm_cols:
                return original_cols[norm_cols.index(sub)]
            raise ValueError(f"Required column '{sub}' must match exactly (case-insensitive).")

        try:
            cct_col = find_exact_col("CCT")
            lux_col = find_exact_col("LUX")
            red_col = find_exact_col("RED")
            green_col = find_exact_col("GREEN")
            blue_col = find_exact_col("BLUE")
            clear_col = find_exact_col("CLEAR")
            wb_col = find_exact_col("WB")
        except ValueError as e:
            self.console_output.append(f"Error: {e}")
            return

        tint = ((sample_time + 1) * (samples + 1)) / 720
        self.console_output.append(f"Calculated TINT: {tint:.4f}")

        # Enforce strictly evaluated matching
        filtered_rows = []
        for index, row in df.iterrows():
            cct = row[cct_col]
            lux = row[lux_col]
            if pd.notna(cct) and pd.notna(lux) and int(cct) in self.target_lux_map:
                if any(np.isclose(float(lux), target_lux) for target_lux in self.target_lux_map[int(cct)]):
                    filtered_rows.append(row)

        if not filtered_rows:
            self.console_output.append("No matching records found based on CCT and LUX mapping.")
            return

        f_df = pd.DataFrame(filtered_rows)
        self.console_output.append(f"Found {len(f_df)} matching raw measurement rows.")

        # Data Reduction: Group by CCT and LUX and take the median
        # to ensure each CCT and LUX combination has exactly one representative row
        f_df = f_df.groupby([cct_col, lux_col], as_index=False).median()
        self.console_output.append(f"After Data Reduction (Median): {len(f_df)} representative rows.")

        scale = (tint * gain) / 256.0
        self.console_output.append(f"Normalisation scale factor: {scale:.4f}")

        # Normalise channels
        for ch in [red_col, green_col, blue_col, clear_col, wb_col]:
            f_df[ch] = f_df[ch] / scale

        Y = f_df[lux_col].values
        X = f_df[[red_col, green_col, blue_col, clear_col, wb_col]].values
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
