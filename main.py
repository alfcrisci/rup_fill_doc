import sys
from PyQt6.QtWidgets import QApplication
from excel_reader_window import ExcelReaderWindow


def main():
    app = QApplication(sys.argv)
    window = ExcelReaderWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()