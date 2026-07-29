import sys
import os
import pandas as pd
import numpy as np
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QTableWidget, QTableWidgetItem, QLineEdit, QLabel,
    QFileDialog, QFormLayout, QGroupBox, QMessageBox
)
from PySide6.QtCore import Qt

class ALS_TestReportParserApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ALS Test Report Parser")
        self.resize(900, 650)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)

        # 1. Directory Selection Layout
        self.dir_layout = QHBoxLayout()
        self.dir_label = QLabel("Directory (目錄):")
        self.dir_input = QLineEdit()
        self.dir_input.setPlaceholderText("Select target directory...")

        # Default directory check: look for "input_files" directory
        default_dir = os.path.join(os.getcwd(), "input_files")
        if os.path.exists(default_dir) and os.path.isdir(default_dir):
            self.dir_input.setText(default_dir)
            self.target_dir = default_dir
        else:
            self.target_dir = ""

        self.dir_btn = QPushButton("Set Directory (設定目錄)")
        self.dir_btn.clicked.connect(self.select_directory)

        self.dir_layout.addWidget(self.dir_label)
        self.dir_layout.addWidget(self.dir_input)
        self.dir_layout.addWidget(self.dir_btn)
        self.main_layout.addLayout(self.dir_layout)

        # 2. Metadata Display Group
        self.meta_group = QGroupBox("Metadata Info (中繼資料)")
        self.meta_layout = QFormLayout(self.meta_group)

        self.dut_version_display = QLineEdit()
        self.dut_version_display.setReadOnly(True)

        self.lens_color_display = QLineEdit()
        self.lens_color_display.setReadOnly(True)

        self.total_rows_display = QLineEdit()
        self.total_rows_display.setReadOnly(True)

        self.total_files_display = QLineEdit()
        self.total_files_display.setReadOnly(True)

        self.meta_layout.addRow("DUT_VERSION:", self.dut_version_display)
        self.meta_layout.addRow("LENS_Color:", self.lens_color_display)
        self.meta_layout.addRow("Total Row Count (總筆數):", self.total_rows_display)
        self.meta_layout.addRow("Total File Count (總檔案數):", self.total_files_display)
        self.main_layout.addWidget(self.meta_group)

        # 3. Action Buttons
        self.run_btn = QPushButton("Run Parsing (執行)")
        self.run_btn.clicked.connect(self.run_parsing)
        self.run_btn.setStyleSheet("font-weight: bold; background-color: #d1e7dd; height: 35px;")
        self.main_layout.addWidget(self.run_btn)

        # 4. Results Table
        self.results_group = QGroupBox("Processed Results (分析結果)")
        self.results_layout = QVBoxLayout(self.results_group)

        self.results_table = QTableWidget()
        self.results_table.setColumnCount(4)
        self.results_table.setHorizontalHeaderLabels([
            "B Column Name (B欄名稱)",
            "Min (最小值)",
            "Max (最大值)",
            "Avg (平均值)"
        ])
        self.results_table.setColumnWidth(0, 300)
        self.results_table.horizontalHeader().setStretchLastSection(True)
        self.results_layout.addWidget(self.results_table)

        self.main_layout.addWidget(self.results_group)

    def select_directory(self):
        dir_path = QFileDialog.getExistingDirectory(self, "Select Test Report Directory")
        if dir_path:
            self.target_dir = dir_path
            self.dir_input.setText(dir_path)

    def run_parsing(self):
        directory = self.dir_input.text().strip()
        if not directory or not os.path.exists(directory):
            QMessageBox.warning(self, "Warning", "Please set a valid directory first.")
            return

        # Find all excel and csv files in the directory
        valid_extensions = (".xlsx", ".xls", ".csv")
        all_files = [
            os.path.join(directory, f) for f in os.listdir(directory)
            if f.lower().endswith(valid_extensions)
        ]

        if not all_files:
            QMessageBox.information(self, "No Files", "No Excel or CSV files found in the specified directory.")
            return

        dut_versions = set()
        lens_colors = set()
        total_valid_rows = 0
        processed_files_count = 0
        b2_header_name = "measurement" # default name
        data_by_label = {} # label -> list of floats

        import csv

        for file_path in all_files:
            rows_data = []
            try:
                if file_path.lower().endswith(".csv"):
                    with open(file_path, "r", encoding="utf-8", errors="replace") as f:
                        reader = csv.reader(f)
                        for r in reader:
                            if len(r) < 3:
                                r = r + [""] * (3 - len(r))
                            rows_data.append(r)
                else:
                    df = pd.read_excel(file_path, header=None)
                    for _, row in df.iterrows():
                        r = []
                        for col_idx in range(max(3, df.shape[1])):
                            if col_idx < df.shape[1]:
                                val = row[col_idx]
                                r.append(str(val).strip() if pd.notna(val) else "")
                            else:
                                r.append("")
                        rows_data.append(r)
            except Exception as e:
                print(f"Skipping file {file_path} due to error: {e}")
                continue

            processed_files_count += 1

            # Extract cell B2 (idx 1 in rows_data, column idx 1) if available
            if len(rows_data) > 1 and len(rows_data[1]) > 1:
                extracted_b2 = str(rows_data[1][1]).strip().strip('"').strip("'")
                if extracted_b2:
                    b2_header_name = extracted_b2

            for idx, r in enumerate(rows_data):
                val_b = str(r[1]).strip()
                val_c_raw = r[2].strip() if isinstance(r[2], str) else r[2]
                val_c = val_c_raw if val_c_raw != "" else None

                # 1. Get lens_color_number (typically at B3 -> C3)
                if val_b == "lens_color_number" and val_c is not None:
                    # Strip quotes if any
                    clean_val = str(val_c).strip().strip('"').strip("'")
                    lens_colors.add(clean_val)

                # 2. Get DUT_VERSION (typically at B4 -> C4)
                elif val_b == "DUT_VERSION" and val_c is not None:
                    clean_val = str(val_c).strip().strip('"').strip("'")
                    dut_versions.add(clean_val)

                # 3. Handle rows from B7 to B732 (0-indexed 6 to 731)
                elif 6 <= idx <= 731:
                    if val_b == "" or val_b == "measurement":
                        continue

                    # Discard rows containing "<Sample_Number>"
                    if "<Sample_Number>" in val_b:
                        continue

                    if val_c is not None:
                        try:
                            # Remove potential quotes around number strings
                            num_str = str(val_c).strip().strip('"').strip("'")
                            num_val = float(num_str)
                            if val_b not in data_by_label:
                                data_by_label[val_b] = []
                            data_by_label[val_b].append(num_val)
                            total_valid_rows += 1
                        except ValueError:
                            pass

        # Update Metadata displays
        self.dut_version_display.setText(", ".join(sorted(list(dut_versions))) if dut_versions else "N/A")
        self.lens_color_display.setText(", ".join(sorted(list(lens_colors))) if lens_colors else "N/A")
        self.total_rows_display.setText(str(total_valid_rows))
        self.total_files_display.setText(str(processed_files_count))

        # Dynamically rename column 0 header label using B2 content (measurement)
        self.results_table.setHorizontalHeaderLabels([
            f"{b2_header_name} (B欄名稱)",
            "Min (最小值)",
            "Max (最大值)",
            "Avg (平均值)"
        ])

        # Populate Results Table in insertion order (original CSV sequence)
        self.results_table.setRowCount(len(data_by_label))
        for row_idx, (label, vals) in enumerate(data_by_label.items()):
            max_val = max(vals)
            min_val = min(vals)
            avg_val = np.mean(vals)

            self.results_table.setItem(row_idx, 0, QTableWidgetItem(label))
            self.results_table.setItem(row_idx, 1, QTableWidgetItem(f"{min_val:.4f}"))
            self.results_table.setItem(row_idx, 2, QTableWidgetItem(f"{max_val:.4f}"))
            self.results_table.setItem(row_idx, 3, QTableWidgetItem(f"{avg_val:.4f}"))

        QMessageBox.information(self, "Success", "Parsing completed successfully!")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ALS_TestReportParserApp()
    window.show()
    sys.exit(app.exec())
