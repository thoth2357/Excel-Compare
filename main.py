import pandas as pd
import sys
import socket
from datetime import datetime
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QDialog,
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
import psycopg2
from argon2 import PasswordHasher
from argon2.exceptions import VerifyMismatchError

# Constants for database connection
DB_HOST = "dpg-cpau3mm3e1ms73a00mug-a.oregon-postgres.render.com"
DB_NAME = "excelcompare"
DB_USER = "thoth"
DB_PASSWORD = "mWxyGn3ykRlm732OyuRDKM1RaAgPllNW"


class LoginDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWindowTitle("Login")
        self.setGeometry(100, 100, 300, 150)

        layout = QVBoxLayout()

        self.username_label = QLabel("Username:")
        layout.addWidget(self.username_label)

        self.username_edit = QLineEdit()
        layout.addWidget(self.username_edit)

        self.pin_label = QLabel("PIN:")
        layout.addWidget(self.pin_label)

        self.pin_edit = QLineEdit()
        self.pin_edit.setEchoMode(QLineEdit.Password)
        layout.addWidget(self.pin_edit)

        self.login_button = QPushButton("Login")
        self.login_button.clicked.connect(self.verify_credentials)
        layout.addWidget(self.login_button)

        self.setLayout(layout)

        self.ph = PasswordHasher()

    def verify_credentials(self):
        username = self.username_edit.text()
        pin = self.pin_edit.text()

        try:
            self.conn = psycopg2.connect(
                host=DB_HOST, dbname=DB_NAME, user=DB_USER, password=DB_PASSWORD
            )
            self.cur = self.conn.cursor()

            statment = f"SELECT user_role, password FROM public.user WHERE username = '{username}'"  # noqa
            self.cur.execute(statment)
            result = self.cur.fetchone()
            if result:
                user_role, hashed_pin = result
                try:
                    self.ph.verify(hashed_pin, pin)
                    self.conn.commit()
                    self.cur.close()
                    self.conn.close()
                    self.accept()
                except VerifyMismatchError:
                    QMessageBox.warning(self, "Error", "Invalid PIN")
            else:
                QMessageBox.warning(self, "Error", "Invalid Username")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))


class ExcelComparator(QMainWindow):
    def __init__(self, username):
        super().__init__()

        self.username = username
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
        filename, _ = file_dialog.getOpenFileName(
            self,
            "Open Original Sheet",
            "",
            "CSV Files (*.csv);;Excel Files (*.xls *.xlsx)",
        )
        if filename:
            self.original_sheet = filename
            self.original_line_edit.setText(self.original_sheet)

            # Populate unique row dropdown
            try:
                if filename.endswith(".csv"):
                    original_data = pd.read_csv(self.original_sheet)
                else:
                    original_data = pd.read_excel(self.original_sheet)
                self.unique_row_dropdown.clear()
                self.unique_row_dropdown.addItems(original_data.columns)
            except Exception as e:
                QMessageBox.critical(
                    self, "Error", f"Failed to load the file: {str(e)}"
                )

    def load_compare_sheet(self):
        file_dialog = QFileDialog()
        filename, _ = file_dialog.getOpenFileName(
            self,
            "Open Compare Sheet",
            "",
            "CSV Files (*.csv);;Excel Files (*.xls *.xlsx)",
        )
        if filename:
            self.compare_sheet = filename
            self.compare_line_edit.setText(self.compare_sheet)

    def compare_sheets(self):
        self.conn = psycopg2.connect(
            host=DB_HOST, dbname=DB_NAME, user=DB_USER, password=DB_PASSWORD
        )
        self.cur = self.conn.cursor()
        if self.original_sheet is None or self.compare_sheet is None:
            return

        try:
            if self.original_sheet.endswith(".csv"):
                original_data = pd.read_csv(self.original_sheet)
            else:
                original_data = pd.read_excel(self.original_sheet)

            if self.compare_sheet.endswith(".csv"):
                compare_data = pd.read_csv(self.compare_sheet)
            else:
                compare_data = pd.read_excel(self.compare_sheet)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load the files: {str(e)}")
            return

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
                        f"<li>Row {row+2} in original sheet is missing in compare sheet. "
                        f"Row Details: {original_row.to_dict()}</li>"
                    )
                if row < len(compare_data):
                    compare_row = compare_data.iloc[row]
                    differences.append(
                        f"<li>Row {row+2} in compare sheet is missing in original sheet. "
                        f"Row Details: {compare_row.to_dict()}</li>"
                    )
                continue

            original_row = original_data.iloc[row]
            compare_row = compare_data.iloc[row]

            if compare_row.isnull().all():
                continue

            row_differences = []

            # Handle fuzzy search if enabled
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
                            account_number = original_row["account no"]
                            difference = f"Row: {row+2}, account no: {account_number}, Column: {original_data.columns[col]} - Values differ"
                            row_differences.append(difference)
            else:
                for col in range(len(original_row)):
                    original_value = original_row.iloc[col]
                    compare_value = compare_row.iloc[col]

                    # Normalize the values
                    if pd.isnull(original_value) or pd.isnull(compare_value):
                        if pd.isnull(original_value) and pd.isnull(compare_value):
                            continue  # both are NaN, consider them equal
                        else:
                            difference = f"Row: {row+2}, Column: {original_data.columns[col]} - Original: {original_value}, Compare: {compare_value}"
                            row_differences.append(difference)
                            continue

                    # Handle numeric comparison with tolerance for floating-point numbers
                    try:
                        original_value_float = float(original_value)
                        compare_value_float = float(compare_value)
                        if (
                            abs(original_value_float - compare_value_float) < 1e-9
                        ):  # Tolerance for floating-point comparison
                            continue  # They are considered equal
                    except ValueError:
                        pass  # One of the values is not a number, fallback to direct comparison

                    # Strip whitespace for string values
                    if isinstance(original_value, str):
                        original_value = original_value.strip()
                    if isinstance(compare_value, str):
                        compare_value = compare_value.strip()

                    # Convert Timestamp to datetime
                    if isinstance(original_value, pd.Timestamp):
                        original_value = original_value.to_pydatetime()
                    if isinstance(compare_value, pd.Timestamp):
                        compare_value = compare_value.to_pydatetime()

                    if original_value != compare_value:
                        difference = f"Row: {row+2}, Column: {original_data.columns[col]} - Original: {original_value}, Compare: {compare_value}"
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
            self.cur.execute(
                "INSERT INTO log (username, log_time, document_differences, interface) VALUES (%s, %s, %s, %s)",  # noqa
                (self.username, datetime.now(), self.differences, "Desktop"),
            )
            self.conn.commit()
            self.cur.close()
            self.conn.close()

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
            pdf.cell(200, 10, txt=f"Computer Name: {computer_name}", ln=True, align="C")
            pdf.cell(200, 10, txt=f"Generated by: {self.username}", ln=True, align="C")
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
                formatted_text += f"{line}\n"
        return formatted_text.strip()

    def show_credits(self):
        credits = "Developer: Oyewunmi Oluwaseyi\nSponsor: Stephen Onuoha Bamidele"
        QMessageBox.information(self, "Credits", credits)


if __name__ == "__main__":
    app = QApplication(sys.argv)

    login = LoginDialog()
    if login.exec() == QDialog.Accepted:
        username = login.username_edit.text()
        window = ExcelComparator(username)
        window.show()
        sys.exit(app.exec())
    else:
        sys.exit()
