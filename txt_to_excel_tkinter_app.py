"""
JSON -> Excel Converter (PyQt5, Dark Theme with Styled Buttons)

Compatible: Python 3.10.11
Dependencies:
    pip install pyqt5 pandas openpyxl
""" 

import sys
import json
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QFileDialog, QMessageBox, QTextEdit
)
from PyQt5.QtGui import QPalette, QColor


def parse_json_from_text(text: str) -> list[dict]:
    text = text.strip("\ufeff\n \r\t")
    if not text:
        raise ValueError("Empty file")

    try:
        obj = json.loads(text)
    except json.JSONDecodeError:
        # maybe NDJSON
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        parsed = [json.loads(ln) for ln in lines]
        return parsed

    if isinstance(obj, list):
        return [item if isinstance(item, dict) else {"value": item} for item in obj]

    if isinstance(obj, dict):
        for key in ("data", "items", "rows", "results", "records"):
            if key in obj and isinstance(obj[key], list):
                return [item if isinstance(item, dict) else {"value": item} for item in obj[key]]
        return [obj]

    raise ValueError("Unknown JSON structure")


def dataframe_from_json_rows(rows: list[dict]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame()
    try:
        df = pd.json_normalize(rows, sep=".")
    except Exception:
        df = pd.DataFrame(rows)
    return df


class ConverterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("JSON → Excel Converter")
        self.setGeometry(300, 100, 800, 600)

        self.loaded_text = None

        layout = QVBoxLayout()

        # Top bar
        top_layout = QHBoxLayout()

        self.btn_open = QPushButton("Import [.txt/.json]")
        self.btn_open.clicked.connect(self.open_file)
        top_layout.addWidget(self.btn_open)

        self.btn_convert = QPushButton("Convert→Save")
        self.btn_convert.setEnabled(False)
        self.btn_convert.clicked.connect(self.convert_and_save)
        top_layout.addWidget(self.btn_convert)

        self.btn_clear = QPushButton("Clear")
        self.btn_clear.clicked.connect(self.clear_text)
        top_layout.addWidget(self.btn_clear)

        # Spacer
        top_layout.addStretch()

        # Info button (circle)
        self.btn_info = QPushButton("ℹ")
        self.btn_info.setFixedSize(40, 40)
        self.btn_info.setStyleSheet("""
            QPushButton {
                border-radius: 20px;
                background-color: #444;
                color: white;
                font-weight: bold;
                font-size: 18px;
            }
            QPushButton:hover {
                background-color: #666;
            }
        """)
        self.btn_info.clicked.connect(self.show_info)
        top_layout.addWidget(self.btn_info)

        layout.addLayout(top_layout)

        # Status
        self.status = QLabel("No file loaded")
        layout.addWidget(self.status)

        # Preview text
        self.preview = QTextEdit()
        self.preview.setReadOnly(True)
        layout.addWidget(self.preview)

        self.setLayout(layout)
        self.apply_dark_theme()
        self.style_buttons()

    def apply_dark_theme(self):
        palette = QPalette()
        palette.setColor(QPalette.Window, QColor(45, 45, 45))
        palette.setColor(QPalette.WindowText, QColor(220, 220, 220))
        palette.setColor(QPalette.Base, QColor(30, 30, 30))
        palette.setColor(QPalette.AlternateBase, QColor(45, 45, 45))
        palette.setColor(QPalette.Text, QColor(220, 220, 220))
        palette.setColor(QPalette.Button, QColor(70, 70, 70))
        palette.setColor(QPalette.ButtonText, QColor(220, 220, 220))
        palette.setColor(QPalette.Highlight, QColor(100, 100, 255))
        palette.setColor(QPalette.HighlightedText, QColor(0, 0, 0))
        self.setPalette(palette)

    def style_buttons(self):
        btn_style = """
            QPushButton {
                background-color: #555;
                color: white;
                font-size: 14px;
                padding: 10px 20px;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #777;
            }
            QPushButton:disabled {
                background-color: #333;
                color: #888;
            }
        """
        self.btn_open.setStyleSheet(btn_style)
        self.btn_convert.setStyleSheet(btn_style)
        self.btn_clear.setStyleSheet(btn_style)

    def open_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open File", "", "Text/JSON Files (*.txt *.json);;All Files (*)")
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                text = f.read()
            self.loaded_text = text
            self.preview.setPlainText(text[:10000] + ("\n\n... (preview truncated) ..." if len(text) > 10000 else ""))
            self.status.setText(f"Loaded: {path} ({len(text)} bytes)")
            self.btn_convert.setEnabled(True)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to open file:\n{e}")

    def clear_text(self):
        self.preview.clear()
        self.loaded_text = None
        self.btn_convert.setEnabled(False)
        self.status.setText("No file loaded")

    def convert_and_save(self):
        if not self.loaded_text:
            QMessageBox.warning(self, "No file", "Please open a file first")
            return
        try:
            rows = parse_json_from_text(self.loaded_text)
            df = dataframe_from_json_rows(rows)
            if df.empty:
                QMessageBox.warning(self, "No Data", "Parsed JSON produced no tabular data")
                return

            path, _ = QFileDialog.getSaveFileName(self, "Save As", "", "Excel Files (*.xlsx);;CSV Files (*.csv)")
            if not path:
                return

            if path.lower().endswith(".csv"):
                df.to_csv(path, index=False, encoding="utf-8")
            else:
                df.to_excel(path, index=False, engine="openpyxl")

            QMessageBox.information(self, "Success", f"Saved {len(df)} rows to:\n{path}")
            self.status.setText(f"Saved to: {path}")
        except Exception as e:
            QMessageBox.critical(self, "Conversion Failed", f"Error:\n{e}")

    def show_info(self):
        msg = QMessageBox(self)
        msg.setWindowTitle("About")
        msg.setText(
            """
            <b>Company Name:</b> <u>MMC CONVERT</u><br><br>
            <b>Developed By:</b> <u>SANJAY CHOURASIYA</u><br><br>
            <b>© 2025 <u>MMC Convert</u></b>. All rights reserved.
        """)
        msg.setStyleSheet("""
            QMessageBox {
                background-color: #2d2d2d;
                color: white;
                font-size: 13px;
            }
            QPushButton {
                background-color: #555;
                color: white;
                padding: 6px 14px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #777;
            }
        """)
        msg.exec_()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ConverterApp()
    window.show()
    sys.exit(app.exec_())
