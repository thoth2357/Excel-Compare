import pandas as pd
import sys
from datetime import datetime
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QFileDialog,
    QVBoxLayout,
    QHBoxLayout,
    QWidget,
    QLabel,
    QLineEdit,
    QPushButton,
    QToolBar,
    QMessageBox,
    QTextEdit,
    QComboBox,
    QRadioButton,
)
from PySide6.QtGui import QIcon, QFont
from PySide6.QtCore import Qt
from fpdf import FPDF
from bs4 import BeautifulSoup

import socket

class ExcelComparator(QMainWindow):
    def __init__(self):
        super().__init__()

        self.original_sheet = None
        self.compare_sheet = None
        self.differences = ""

        self.setWindowTitle("Excel Comparator")
        self.setGeometry(100, 100, 600, 400)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout(self.central_widget)
        self.layout.setContentsMargins(20, 20, 20, 20)
        self.layout.setSpacing(20)

        self.toolbar = QToolBar()
        self.addToolBar(self.toolbar)

        self.credits_button = QPushButton("Credits")
        self.credits_button.setIcon(QIcon.fromTheme("help-about"))
        self.credits_button.clicked.connect(self.show_credits)
        self.toolbar.addWidget(self.credits_button)

        # Settings Toolbar
        self.settings_toolbar = QToolBar()
        self.addToolBar(Qt.BottomToolBarArea, self.settings_toolbar)

        self.fuzzy_radio_button = QRadioButton("Fuzzy Search")
        self.settings_toolbar.addWidget(self.fuzzy_radio_button)

        self.unique_row_label = QLabel("Unique Row:")
        self.settings_toolbar.addWidget(self.unique_row_label)

        self.unique_row_dropdown = QComboBox()
        self.unique_row_dropdown.setEnabled(False)
        self.settings_toolbar.addWidget(self.unique_row_dropdown)

        self.original_label = QLabel("Original Sheet: ")
        self.layout.addWidget(self.original_label)

        self.original_line_edit = QLineEdit()
        self.original_line_edit.setReadOnly(True)
        self.layout.addWidget(self.original_line_edit)

        original_button_layout = QHBoxLayout()
        self.original_button = QPushButton("Browse")
        self.original_button.setIcon(QIcon.fromTheme("document-open"))
        self.original_button.clicked.connect(self.load_original_sheet)
        original_button_layout.addWidget(self.original_button)

        self.layout.addLayout(original_button_layout)

        self.compare_label = QLabel("Compare Sheet: ")
        self.layout.addWidget(self.compare_label)

        self.compare_line_edit = QLineEdit()
        self.compare_line_edit.setReadOnly(True)
        self.layout.addWidget(self.compare_line_edit)

        compare_button_layout = QHBoxLayout()
        self.compare_button = QPushButton("Browse")
        self.compare_button.setIcon(QIcon.fromTheme("document-open"))
        self.compare_button.clicked.connect(self.load_compare_sheet)
        compare_button_layout.addWidget(self.compare_button)

        self.layout.addLayout(compare_button_layout)

        self.compare_button = QPushButton("Compare")
        self.compare_button.setIcon(QIcon.fromTheme("document-properties"))
        self.compare_button.clicked.connect(self.compare_sheets)
        self.layout.addWidget(self.compare_button)

        self.download_button = QPushButton("Download Report")
        self.download_button.setIcon(QIcon.fromTheme("document-save"))
        self.download_button.clicked.connect(self.download_report)
        self.layout.addWidget(self.download_button)
        self.download_button.setEnabled(False)  # Initially disabled

        self.result_label = QLabel("Differences: ")
        self.layout.addWidget(self.result_label)

        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        self.layout.addWidget(self.result_text)

        self.apply_styles()

        self.fuzzy_radio_button.toggled.connect(self.enable_disable_dropdown)

    def apply_styles(self):
        self.setStyleSheet("""
            QLabel {
                font-size: 14px;
                font-weight: bold;
            }
            QLineEdit {
                padding: 5px;
                font-size: 14px;
            }
            QPushButton {
                padding: 10px;
                font-size: 14px;
                background-color: #6200EE;
                color: white;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #3700B3;
            }
            QRadioButton {
                font-size: 14px;
            }
            QComboBox {
                padding: 5px;
                font-size: 14px;
            }
            QTextEdit {
                background-color: #fafafa;
                padding: 10px;
                font-family: 'Courier New', monospace;
                font-size: 16px;
            }
        """)

        self.result_text.setFont(QFont("Courier New", 16))
        self.result_text.setStyleSheet("background-color: #fafafa; color: #333333;")

    def enable_disable_dropdown(self):
        self.unique_row_dropdown.setEnabled(self.fuzzy_radio_button.isChecked())

    def load_original_sheet(self):
        file_dialog = QFileDialog()
        filename = file_dialog.getOpenFileName(self, "Open Original Sheet")
        self.original_sheet = filename[0]
        self.original_line_edit.setText(self.original_sheet)

        # Populate unique row dropdown
        if self.original_sheet:
            original_data = pd.read_excel(self.original_sheet)
            self.unique_row_dropdown.clear()
            self.unique_row_dropdown.addItems(original_data.columns)

    def load_compare_sheet(self):
        file_dialog = QFileDialog()
        filename = file_dialog.getOpenFileName(self, "Open Compare Sheet")
        self.compare_sheet = filename[0]
        self.compare_line_edit.setText(self.compare_sheet)

    def compare_sheets(self):
        if self.original_sheet is None or self.compare_sheet is None:
            return

        original_data = pd.read_excel(self.original_sheet)
        compare_data = pd.read_excel(self.compare_sheet)

        # Check if columns match
        if not original_data.columns.equals(compare_data.columns):
            self.result_text.setText(
                "Columns in the two sheets do not match. Comparison aborted."
            )
            return

        differences = []

        # Check for additional columns in compare sheet
        additional_columns = set(compare_data.columns) - set(original_data.columns)
        if additional_columns:
            differences.append("<b>Additional columns found in compare sheet:</b><ul>")
            for col in additional_columns:
                differences.append(f"<li>{col}</li>")
            differences.append("</ul>")

        row_differences_header = False

        for row in range(max(len(original_data), len(compare_data))):
            if row >= len(original_data) or row >= len(compare_data):
                if not row_differences_header:
                    differences.append("<b>Row Differences:</b><ul>")
                    row_differences_header = True
                if row < len(original_data):
                    original_row = original_data.iloc[row]
                    differences.append(
                        f"<li>Row {row+1} in original sheet is missing in compare sheet. "
                        f"Row Details: {original_row.to_dict()}</li>"
                    )
                if row < len(compare_data):
                    compare_row = compare_data.iloc[row]
                    differences.append(
                        f"<li>Row {row+1} in compare sheet is missing in original sheet. "
                        f"Row Details: {compare_row.to_dict()}</li>"
                    )
                continue

            original_row = original_data.iloc[row]
            compare_row = compare_data.iloc[row]

            if compare_row.isnull().all():
                continue

            row_differences = []

            if self.fuzzy_radio_button.isChecked():
                unique_row_column = self.unique_row_dropdown.currentText()

                for col in range(len(original_row)):
                    original_value = original_row.iloc[col]
                    compare_value = compare_row.iloc[col]

                    if isinstance(original_value, str) and isinstance(
                        compare_value, str
                    ):
                        original_value = original_value.strip()
                        compare_value = compare_value.strip()

                    if isinstance(original_value, pd.Timestamp):
                        original_value = original_value.to_pydatetime()
                    if isinstance(compare_value, pd.Timestamp):
                        compare_value = compare_value.to_pydatetime()

                    if original_data.columns[col] == unique_row_column:
                        if original_value != compare_value:
                            account_number = original_row["Account No"]
                            difference = f"Row: {row+1}, Account No: {account_number}, Column: {original_data.columns[col]} - Values differ"
                            row_differences.append(difference)
            else:
                for col in range(len(original_row)):
                    original_value = original_row.iloc[col]
                    compare_value = compare_row.iloc[col]

                    if isinstance(original_value, str):
                        original_value = original_value.strip()
                    if isinstance(compare_value, str):
                        compare_value = compare_value.strip()

                    if isinstance(original_value, pd.Timestamp):
                        original_value = original_value.to_pydatetime()
                    if isinstance(compare_value, pd.Timestamp):
                        compare_value = compare_value.to_pydatetime()

                    if original_value != compare_value:
                        difference = f"Row: {row+1}, Column: {original_data.columns[col]} - Original: {original_value}, Compare: {compare_value}"
                        row_differences.append(difference)

            if row_differences:
                if not row_differences_header:
                    differences.append("<b>Row Differences:</b><ul>")
                    row_differences_header = True
                for diff in row_differences:
                    differences.append(f"<li>{diff}</li>")

        if row_differences_header:
            differences.append("</ul>")

        self.differences = "".join(differences)

        if not differences:
            self.result_text.setText("No differences found")
        else:
            self.result_text.setHtml(self.differences)
            self.download_button.setEnabled(
                True
            )  # Enable the download button if differences are found

    def download_report(self):
        if not self.differences:
            QMessageBox.warning(self, "Warning", "No differences to report.")
            return

        save_dialog = QFileDialog()
        file_path, _ = save_dialog.getSaveFileName(
            self, "Save Report", "", "PDF Files (*.pdf)"
        )

        if file_path:
            if not file_path.endswith(".pdf"):
                file_path += ".pdf"
                
            computer_name = socket.gethostname()
    
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)

            pdf.cell(200, 10, txt="Excel Comparison Report", ln=True, align="C")
            pdf.cell(
                200,
                10,
                txt=f"Report Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                ln=True,
                align="C",
            )
            pdf.cell(200, 10, txt=f"Computer Name: {computer_name}", ln=True, align='C')
            pdf.ln(10)

            pdf.set_font("Arial", size=10)
            pdf.multi_cell(0, 10, txt="Original Sheet: " + self.original_sheet)
            pdf.multi_cell(0, 10, txt="Compare Sheet: " + self.compare_sheet)
            pdf.ln(10)

            pdf.set_font("Arial", size=12)
            pdf.multi_cell(0, 10, txt="Differences:", align="L")
            pdf.set_font("Arial", size=10)
            pdf.multi_cell(
                0, 10, txt=self.format_differences_for_pdf(self.differences), align="L"
            )

            pdf.output(file_path)
            QMessageBox.information(self, "Success", "Report saved successfully!")

    def format_differences_for_pdf(self, html):
        """Convert HTML content to formatted plain text for the PDF report."""
        soup = BeautifulSoup(html, "html.parser")
        text = soup.get_text(separator="\n")
        formatted_text = ""
        lines = text.split("\n")
        for line in lines:
            if "Row:" in line:
                formatted_text += f"\n{line}\n"
            else:
                formatted_text += f"\n{line}\n"
        return formatted_text.strip()

    def show_credits(self):
        credits = "Developer: Oyewunmi Oluwaseyi\nSponsor: Stephen Onuoha Bamidele"
        QMessageBox.information(self, "Credits", credits)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelComparator()
    window.show()
    sys.exit(app.exec())
