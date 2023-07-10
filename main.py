import sys
import pandas as pd
from PySide6.QtWidgets import QApplication, QMainWindow, QFileDialog, QVBoxLayout, QWidget, QLabel, QLineEdit, QPushButton, QToolBar, QMessageBox, QTextEdit


class ExcelComparator(QMainWindow):
    def __init__(self):
        super().__init__()

        self.original_sheet = None
        self.compare_sheet = None

        self.setWindowTitle("Excel Comparator")
        self.setGeometry(100, 100, 500, 300)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout(self.central_widget)

        self.toolbar = QToolBar()
        self.addToolBar(self.toolbar)

        self.credits_button = QPushButton("Credits")
        self.credits_button.clicked.connect(self.show_credits)
        self.toolbar.addWidget(self.credits_button)

        self.original_label = QLabel("Original Sheet: ")
        self.layout.addWidget(self.original_label)

        self.original_line_edit = QLineEdit()
        self.original_line_edit.setReadOnly(True)
        self.layout.addWidget(self.original_line_edit)

        self.original_button = QPushButton("Browse")
        self.original_button.clicked.connect(self.load_original_sheet)
        self.layout.addWidget(self.original_button)

        self.compare_label = QLabel("Compare Sheet: ")
        self.layout.addWidget(self.compare_label)

        self.compare_line_edit = QLineEdit()
        self.compare_line_edit.setReadOnly(True)
        self.layout.addWidget(self.compare_line_edit)

        self.compare_button = QPushButton("Browse")
        self.compare_button.clicked.connect(self.load_compare_sheet)
        self.layout.addWidget(self.compare_button)

        self.compare_button = QPushButton("Compare")
        self.compare_button.clicked.connect(self.compare_sheets)
        self.layout.addWidget(self.compare_button)

        self.result_label = QLabel("Differences: ")
        self.layout.addWidget(self.result_label)

        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        self.layout.addWidget(self.result_text)

    def load_original_sheet(self):
        file_dialog = QFileDialog()
        filename = file_dialog.getOpenFileName(self, 'Open Original Sheet')
        self.original_sheet = filename[0]
        self.original_line_edit.setText(self.original_sheet)

    def load_compare_sheet(self):
        file_dialog = QFileDialog()
        filename = file_dialog.getOpenFileName(self, 'Open Compare Sheet')
        self.compare_sheet = filename[0]
        self.compare_line_edit.setText(self.compare_sheet)

    def compare_sheets(self):
        if self.original_sheet is None or self.compare_sheet is None:
            return

        original_data = pd.read_excel(self.original_sheet)
        compare_data = pd.read_excel(self.compare_sheet)

        differences = []

        for row in range(max(len(original_data), len(compare_data))):
            if row >= len(original_data) or row >= len(compare_data):
                difference = f"Row: {row+1} - Rows are mismatched\n"
                differences.append(difference)
            else:
                original_row = original_data.iloc[row]
                compare_row = compare_data.iloc[row]

                if compare_row.isnull().all():
                    continue

                row_differences = []

                for col in range(len(original_row)):
                    if original_row.iloc[col] != compare_row.iloc[col]:
                        difference = f"Row: {row+1}, Column: {col+1} - Values differ\n"
                        row_differences.append(difference)

                if row_differences:
                    differences.extend(row_differences)

        if not differences:
            self.result_text.setText("No differences found")
        else:
            self.result_text.setText("".join(differences))

    def show_credits(self):
        credits = "Developer: Oyewunmi Oluwaseyi\nSponsor: Stephen Onuoha Bamidele"
        QMessageBox.information(self, "Credits", credits)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelComparator()
    window.show()
    sys.exit(app.exec())
