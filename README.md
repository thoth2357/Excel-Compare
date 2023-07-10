# Excel-Compare
ExcelCompare is a desktop application designed as a small tool for banking staff to process people's salaries through Excel sheets. The motivation behind this project is to provide a simple and efficient way for banking staff to check for discrepancies in case someone modifies the Excel sheets in any way.

The application allows users to compare two Excel sheets and display the differences between them. It provides a user-friendly interface for browsing and selecting the original and compare sheets, and it highlights the differences row-wise. The application is built using PySide6 for the graphical user interface and pandas for reading and comparing Excel data.

## Features
- Select and load the original Excel sheet.
- Select and load the compare Excel sheet.
- Compare the two sheets and identify differences.
- Display the differences in a read-only field.

## Prerequisites
- Python 3.x
- PySide6
- pandas
- Poetry (optional, for simplified dependency management)

## Installation
1. Clone the repository:


```bash
git clone https://github.com/your-username/excel-comparator.git
```
2. Change to the project directory:

```bash
cd excel-comparator
```

3. Install the required dependencies using pip:

```bash
pip install PySide6 pandas
```
  Alternatively, if you prefer to use Poetry for dependency management:
```bash
poetry install
```
If you don't have Poetry installed, you can follow the installation instructions from the official Poetry documentation: https://python-poetry.org/docs/#installation

## Usage
1. Run the application:

```bash
Copy code
python main.py
```

2. The Excel Comparator window will appear.

3. Click the "Browse" button next to "Original Sheet" and select the Excel sheet you want to use as the original.

4. Click the "Browse" button next to "Compare Sheet" and select the Excel sheet you want to compare against the original.

5. Click the "Compare" button to start the comparison process.

6. The differences between the original and compare sheets will be displayed in the "Differences" field.

7. Click the "Credits" button to view the developer and sponsor information.

## Contributing
Contributions to ExcelComparator are welcome! If you find a bug or want to suggest an enhancement, please open an issue or submit a pull request.

## License
This project is licensed under the MIT License.


## Acknowledgments
- [PySide6](https://wiki.qt.io/Qt_for_Python)
- [pandas](https://pandas.pydata.org/)

