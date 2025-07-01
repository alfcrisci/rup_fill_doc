# excel_reader_window.py

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
    QLineEdit, QPushButton, QFileDialog, QTextEdit, QMessageBox,
    QListWidget, QListWidgetItem
)
from PyQt6.QtCore import Qt
import pandas as pd
import openpyxl

class ExcelReaderWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("RUP IBE helper")
        self.setGeometry(100, 100, 800, 600)
        
        # Central widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout()
        central_widget.setLayout(layout)
        
        # File selection
        file_layout = QHBoxLayout()
        self.file_path = QLineEdit()
        self.file_path.setPlaceholderText("Seleziona il file della procedura...")
        file_button = QPushButton("Browse...")
        file_button.clicked.connect(self.browse_file)
        file_layout.addWidget(self.file_path)
        file_layout.addWidget(file_button)
        
        # Sheet selection
        sheet_layout = QHBoxLayout()
        sheet_layout.addWidget(QLabel("Nome del foglio:"))
        self.sheet_input = QLineEdit()
        self.sheet_input.setPlaceholderText("Lascia bianco per foglio attivo....")
        sheet_layout.addWidget(self.sheet_input)
        
        # Variable list
        self.variable_list = QListWidget()
        self.variable_list.itemDoubleClicked.connect(self.show_variable_value)
        layout.addWidget(QLabel("Variables found:"))
        layout.addWidget(self.variable_list)
        
        # Results display
        self.results_display = QTextEdit()
        self.results_display.setReadOnly(True)
        
        # Add all widgets to main layout
        layout.addLayout(file_layout)
        layout.addLayout(sheet_layout)
        layout.addWidget(self.results_display)
        
        # Scan button
        scan_button = QPushButton("Scan for Variables")
        scan_button.clicked.connect(self.scan_variables)
        layout.addWidget(scan_button)
    
    def browse_file(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self, "Apri file Excel", "", "Excel Files (*.xlsx *.xls)"
        )
        if file_name:
            self.file_path.setText(file_name)
            self.scan_variables()
    
    def scan_variables(self):
        file_path = self.file_path.text()
        if not file_path:
            QMessageBox.warning(self, "Warning", "Seleziona prima un file excel.")
            return
            
        try:
            # Read the Excel file
            workbook = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            
            # Get the specified sheet or active sheet
            sheet_name = self.sheet_input.text().strip()
            sheet = workbook[sheet_name] if sheet_name else workbook.active
            
            self.variable_list.clear()
            variables_found = False
            
            # Scan for cells with values
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value is not None:
                        # Consider cells with text as potential variables
                        if isinstance(cell.value, str) and cell.value.strip():
                            item = QListWidgetItem(f"{cell.coordinate}: {cell.value}")
                            item.setData(Qt.ItemDataRole.UserRole, cell.coordinate)
                            self.variable_list.addItem(item)
                            variables_found = True
            
            if not variables_found:
                self.results_display.setPlainText("No variables found in the selected sheet.")
            
            workbook.close()
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to read Excel file:\n{str(e)}")
    
    def show_variable_value(self, item):
        file_path = self.file_path.text()
        if not file_path:
            return
            
        cell_ref = item.data(Qt.ItemDataRole.UserRole)
        
        try:
            workbook = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            
            # Get the specified sheet or active sheet
            
            sheet_name = self.sheet_input.text().strip()
            sheet = workbook[sheet_name] if sheet_name else workbook.active
            
            cell = sheet[cell_ref]
            self.results_display.setPlainText(
                f"Variable: {cell.value}\n"
                f"Location: {cell_ref}\n"
                f"Data type: {type(cell.value).__name__}"
            )
            
            workbook.close()
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to read cell:\n{str(e)}")

if __name__ == "__main__":
    from PyQt6.QtWidgets import QApplication
    import sys
    
    app = QApplication(sys.argv)
    window = ExcelReaderWindow()
    window.show()
    sys.exit(app.exec())