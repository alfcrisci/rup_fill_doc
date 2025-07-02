import os
import sys
from datetime import datetime
from docxtpl import DocxTemplate
import openpyxl
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QPushButton, QFileDialog, QTextEdit, QMessageBox,
    QListWidget, QListWidgetItem, QGroupBox, QTabWidget,
    QFormLayout, QSplitter, QFrame
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPixmap, QIcon
from PyQt6.QtWidgets import QApplication

class ExcelReaderWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("RUP IBE helper - Document Generator")
        self.setGeometry(100, 100, 1200, 850)
        
        # Set application icon (replace with your icon if needed)
        self.setWindowIcon(QIcon(":/images/app_icon.png"))
        
        # Central widget with tabs
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)
        
        # =============================================
        # HEADER SECTION (Customizable Title, Logo, Credits)
        # =============================================
        header_frame = QFrame()
        header_frame.setFrameShape(QFrame.Shape.StyledPanel)
        header_frame.setStyleSheet("background-color: #f8f9fa; border-bottom: 1px solid #dee2e6;")
        header_layout = QHBoxLayout()
        header_frame.setLayout(header_layout)
        
        # Left section for logo
        self.logo_label = QLabel()
        self.logo_label.setFixedSize(320, 120)
        self.logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.logo_label.setStyleSheet("border: none;")
        header_layout.addWidget(self.logo_label)
        
        # Center section for title
        title_frame = QFrame()
        title_layout = QVBoxLayout()
        title_frame.setLayout(title_layout)
        
        self.title_label = QLabel("RUP IBE HELPER")
        self.title_label.setStyleSheet("""
            font-size: 28px; 
            font-weight: bold; 
            color: #2c3e50;
            margin-bottom: 5px;
        """)
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.subtitle_label = QLabel("Generatore di documenti per procedure d'acquisto")
        self.subtitle_label.setStyleSheet("""
            font-size: 14px; 
            color: #7f8c8d;
            font-style: italic;
        """)
        self.subtitle_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        title_layout.addWidget(self.title_label)
        title_layout.addWidget(self.subtitle_label)
        header_layout.addWidget(title_frame, stretch=1)
        
        # Right section for credits
        credits_frame = QFrame()
        credits_layout = QVBoxLayout()
        credits_frame.setLayout(credits_layout)
        
        self.version_label = QLabel("Versione 1.0")
        self.author_label = QLabel("Sviluppato da: IBE CNR")
        self.date_label = QLabel(datetime.now().strftime("%d/%m/%Y"))
        
        for label in [self.version_label, self.author_label, self.date_label]:
            label.setStyleSheet("""
                font-size: 12px; 
                color: #7f8c8d;
                margin: 2px;
            """)
            label.setAlignment(Qt.AlignmentFlag.AlignRight)
            credits_layout.addWidget(label)
        
        header_layout.addWidget(credits_frame)
        
        # Add header to main layout
        main_layout.addWidget(header_frame)
        
        # Separator line
        separator = QFrame()
        separator.setFrameShape(QFrame.Shape.HLine)
        separator.setFrameShadow(QFrame.Shadow.Sunken)
        separator.setStyleSheet("margin: 5px 0; border-color: #dee2e6;")
        main_layout.addWidget(separator)
        
        # =============================================
        # MAIN CONTENT (Tabs)
        # =============================================
        # Create tabs
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
            QTabBar::tab {
                padding: 8px;
                min-width: 120px;
            }
            QTabBar::tab:selected {
                background: #e9ecef;
                border-bottom: 2px solid #2c3e50;
            }
        """)
        main_layout.addWidget(self.tabs)
        
        # Excel Reader Tab
        self.setup_excel_tab()
        
        # Document Generator Tab
        self.setup_document_tab()
        
        # Current data storage
        self.current_file = None
        self.sheet_data = {
            'dati_generali_procedura': None,
            'generazioni_offerte': None
        }
        
        # Load default logo (replace with your logo path)
        self.load_logo("images/default_logo.png")

    def load_logo(self, path):
        """Load logo image with fallback to text"""
        try:
            pixmap = QPixmap(path)
            if not pixmap.isNull():
                self.logo_label.setPixmap(pixmap.scaled(
                    self.logo_label.size(),
                    Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.SmoothTransformation
                ))
                return
        except:
            pass
        
        # Fallback text logo
        self.logo_label.setText("LOGO")
        self.logo_label.setStyleSheet("""
            background-color: #e9ecef;
            border-radius: 4px;
            font-weight: bold;
            font-size: 24px;
            color: #2c3e50;
        """)

    def setup_excel_tab(self):
        """Setup the Excel reader tab"""
        excel_tab = QWidget()
        excel_layout = QVBoxLayout()
        excel_tab.setLayout(excel_layout)
        
        # File selection
        file_group = QGroupBox("Selezione File Excel")
        file_group.setStyleSheet("QGroupBox { font-weight: bold; }")
        file_layout = QHBoxLayout()
        self.file_path = QLineEdit()
        self.file_path.setPlaceholderText("Seleziona file Excel...")
        file_button = QPushButton("Sfoglia...")
        file_button.setIcon(QIcon.fromTheme("document-open"))
        file_button.clicked.connect(self.browse_excel_file)
        file_layout.addWidget(self.file_path)
        file_layout.addWidget(file_button)
        file_group.setLayout(file_layout)
        
        # Splitter for the two lists
        splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # List for dati_generali_procedura
        self.dati_generali_list = QListWidget()
        self.dati_generali_list.itemDoubleClicked.connect(
            lambda item: self.show_variable_value(item, 'dati_generali_procedura'))
        
        # List for generazioni_offerte
        self.generazioni_offerte_list = QListWidget()
        self.generazioni_offerte_list.itemDoubleClicked.connect(
            lambda item: self.show_variable_value(item, 'generazioni_offerte'))
        
        # Add lists to splitter
        splitter.addWidget(self.create_list_group(self.dati_generali_list, "Dati Generali Procedura"))
        splitter.addWidget(self.create_list_group(self.generazioni_offerte_list, "Generazioni Offerte"))
        
        # Results display
        self.results_display = QTextEdit()
        self.results_display.setReadOnly(True)
        self.results_display.setStyleSheet("""
            QTextEdit {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                padding: 5px;
            }
        """)
        
        # Add widgets to Excel tab
        excel_layout.addWidget(file_group)
        excel_layout.addWidget(splitter)
        excel_layout.addWidget(QLabel("Output:"))
        excel_layout.addWidget(self.results_display)
        
        # Scan button
        scan_button = QPushButton("Leggi Fogli Excel")
        scan_button.setIcon(QIcon.fromTheme("document-open"))
        scan_button.setStyleSheet("""
            QPushButton {
                padding: 5px;
                background-color: #2c3e50;
                color: white;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #34495e;
            }
        """)
        scan_button.clicked.connect(self.read_excel_sheets)
        excel_layout.addWidget(scan_button)
        
        self.tabs.addTab(excel_tab, "Excel Reader")

    def create_list_group(self, list_widget, title):
        """Create a group box for a list widget"""
        group = QGroupBox(title)
        group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                margin-top: 10px;
            }
        """)
        layout = QVBoxLayout()
        
        list_widget.setStyleSheet("""
            QListWidget {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                padding: 5px;
            }
        """)
        
        layout.addWidget(list_widget)
        group.setLayout(layout)
        return group

    def setup_document_tab(self):
        """Setup the Document Generator tab"""
        doc_tab = QWidget()
        doc_layout = QVBoxLayout()
        doc_tab.setLayout(doc_layout)
        
        # Template selection
        template_group = QGroupBox("Configurazione Documento")
        template_group.setStyleSheet("QGroupBox { font-weight: bold; }")
        template_layout = QFormLayout()
        
        self.template_path = QLineEdit()
        self.template_path.setPlaceholderText("Seleziona template Word...")
        template_btn = QPushButton("Sfoglia...")
        template_btn.setIcon(QIcon.fromTheme("document-open"))
        template_btn.clicked.connect(self.browse_template_file)
        
        self.output_dir = QLineEdit()
        self.output_dir.setPlaceholderText("Cartella di output...")
        output_btn = QPushButton("Sfoglia...")
        output_btn.setIcon(QIcon.fromTheme("folder-open"))
        output_btn.clicked.connect(self.browse_output_dir)
        
        template_layout.addRow("Template Word:", self.template_path)
        template_layout.addRow(template_btn)
        template_layout.addRow("Cartella Output:", self.output_dir)
        template_layout.addRow(output_btn)
        
        # Document fields
        self.doc_fields = {
            'numero_CUP': QLineEdit(),
            'servizio_fornitura': QLineEdit(),
            'prestazione_servizio_fornitura': QLineEdit(),
            'nome_cognome': QLineEdit(),
            'mail_contatto': QLineEdit(),
            'acronimo_progetto': QLineEdit(),
            'oggetto_fornitura_servizio': QLineEdit(),
            'nome_ditta': QLineEdit(),
            'indirizzo_ditta': QLineEdit(),
            'cap_ditta': QLineEdit(),
            'pec_ditta': QLineEdit()
        }
        
        for field_name, field_widget in self.doc_fields.items():
            label = field_name.replace('_', ' ').title() + ":"
            field_widget.setPlaceholderText(f"Inserisci {label.lower()}")
            template_layout.addRow(label, field_widget)
        
        template_group.setLayout(template_layout)
        
        # Generate buttons
        button_layout = QHBoxLayout()
        
        generate_dati_btn = QPushButton("Genera da Dati Generali")
        generate_dati_btn.setIcon(QIcon.fromTheme("document-save-as"))
        generate_dati_btn.setStyleSheet("""
            QPushButton {
                padding: 8px;
                background-color: #27ae60;
                color: white;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #2ecc71;
            }
        """)
        generate_dati_btn.clicked.connect(lambda: self.generate_document('dati_generali_procedura'))
        
        generate_offerte_btn = QPushButton("Genera da Offerte")
        generate_offerte_btn.setIcon(QIcon.fromTheme("document-save-as"))
        generate_offerte_btn.setStyleSheet("""
            QPushButton {
                padding: 8px;
                background-color: #2980b9;
                color: white;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #3498db;
            }
        """)
        generate_offerte_btn.clicked.connect(lambda: self.generate_document('generazioni_offerte'))
        
        button_layout.addWidget(generate_dati_btn)
        button_layout.addWidget(generate_offerte_btn)
        
        # Add widgets to Document tab
        doc_layout.addWidget(template_group)
        doc_layout.addLayout(button_layout)
        doc_layout.addStretch()
        
        self.tabs.addTab(doc_tab, "Document Generator")

    # Excel Reader functions
    def browse_excel_file(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self, "Apri file Excel", "", "Excel Files (*.xlsx *.xls)"
        )
        if file_name:
            self.file_path.setText(file_name)
            self.current_file = file_name
            self.results_display.append(f"File selezionato: {file_name}")

    def read_excel_sheets(self):
        if not self.current_file:
            QMessageBox.warning(self, "Attenzione", "Seleziona prima un file excel.")
            return
            
        try:
            workbook = openpyxl.load_workbook(self.current_file, read_only=True, data_only=True)
            
            # Clear previous data
            self.dati_generali_list.clear()
            self.generazioni_offerte_list.clear()
            self.sheet_data = {
                'dati_generali_procedura': None,
                'generazioni_offerte': None
            }
            
            # Read specified sheets
            for sheet_name in self.sheet_data.keys():
                try:
                    sheet = workbook[sheet_name]
                    self.sheet_data[sheet_name] = sheet
                    
                    # Get the appropriate list widget
                    target_list = (self.dati_generali_list if sheet_name == 'dati_generali_procedura' 
                                 else self.generazioni_offerte_list)
                    
                    # Scan for cells with values
                    for row in sheet.iter_rows():
                        for cell in row:
                            if cell.value is not None and isinstance(cell.value, str) and cell.value.strip():
                                item = QListWidgetItem(f"{cell.coordinate}: {cell.value}")
                                item.setData(Qt.ItemDataRole.UserRole, cell.coordinate)
                                target_list.addItem(item)
                    
                    self.results_display.append(f"Trovate {target_list.count()} variabili nel foglio {sheet_name}")
                
                except KeyError:
                    QMessageBox.warning(self, "Attenzione", f"Il foglio '{sheet_name}' non esiste nel file.")
                    continue
            
            workbook.close()
            
        except Exception as e:
            QMessageBox.critical(self, "Errore", f"Errore nella lettura del file Excel:\n{str(e)}")

    def show_variable_value(self, item, sheet_name):
        if not self.current_file or sheet_name not in self.sheet_data or not self.sheet_data[sheet_name]:
            return
            
        cell_ref = item.data(Qt.ItemDataRole.UserRole)
        
        try:
            sheet = self.sheet_data[sheet_name]
            cell = sheet[cell_ref]
            
            self.results_display.append(
                f"\nDettaglio variabile:\n"
                f"• Foglio: {sheet_name}\n"
                f"• Posizione: {cell_ref}\n"
                f"• Valore: {cell.value}\n"
                f"• Tipo: {type(cell.value).__name__}\n"
                f"----------------------------")
            
        except Exception as e:
            QMessageBox.critical(self, "Errore", f"Errore nella lettura della cella:\n{str(e)}")

    # Document Generator functions
    def browse_template_file(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self, "Seleziona template Word", "", "Word Documents (*.docx)"
        )
        if file_name:
            self.template_path.setText(file_name)

    def browse_output_dir(self):
        dir_name = QFileDialog.getExistingDirectory(
            self, "Seleziona cartella di output"
        )
        if dir_name:
            self.output_dir.setText(dir_name)

    def generate_document(self, source_sheet):
        """Generate the Word document from template using data from specified sheet"""
        template_path = self.template_path.text()
        output_dir = self.output_dir.text()
        
        if not template_path:
            QMessageBox.warning(self, "Attenzione", "Seleziona un template Word.")
            return
        
        if not output_dir:
            output_dir = "A_preventivo"
            self.output_dir.setText(output_dir)
        
        if source_sheet not in self.sheet_data or not self.sheet_data[source_sheet]:
            QMessageBox.warning(self, "Attenzione", f"Nessun dato disponibile dal foglio {source_sheet}")
            return
        
        try:
            # Create output directory if not exists
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # Prepare context
            context = {
                'numero_CUP': self.doc_fields['numero_CUP'].text(),
                'servizio_fornitura': self.doc_fields['servizio_fornitura'].text(),
                'prestazione_servizio_fornitura': self.doc_fields['prestazione_servizio_fornitura'].text(),
                'nome_cognome': self.doc_fields['nome_cognome'].text(),
                'mail_contatto': self.doc_fields['mail_contatto'].text(),
                'acronimo_progetto': self.doc_fields['acronimo_progetto'].text(),
                'oggetto_fornitura_servizio': self.doc_fields['oggetto_fornitura_servizio'].text(),
                'nome_ditta': self.doc_fields['nome_ditta'].text(),
                'indirizzo_ditta': self.doc_fields['indirizzo_ditta'].text(),
                'cap_ditta': self.doc_fields['cap_ditta'].text(),
                'pec_ditta': self.doc_fields['pec_ditta'].text(),
                'data_corrente': datetime.now().strftime('%d/%m/%Y')
            }
            
            # Add data from the selected sheet
            sheet = self.sheet_data[source_sheet]
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value is not None and isinstance(cell.value, str) and cell.value.strip():
                        context[f"{source_sheet}_{cell.coordinate}"] = cell.value
            
            # Generate filename
            cognome = context['nome_cognome'].split()[-1] if context['nome_cognome'] else 'documento'
            source_suffix = "DatiGenerali" if source_sheet == 'dati_generali_procedura' else "Offerte"
            output_filename = f"RichiestaOfferta_{cognome}_{source_suffix}.docx"
            output_path = os.path.join(output_dir, output_filename)
            
            # Render and save document
            doc = DocxTemplate(template_path)
            doc.render(context)
            doc.save(output_path)
            
            QMessageBox.information(
                self, 
                "Successo", 
                f"Documento generato con successo da {source_sheet}!\nSalvato in:\n{output_path}"
            )
            
        except Exception as e:
            QMessageBox.critical(
                self, 
                "Errore", 
                f"Errore durante la generazione del documento:\n{str(e)}"
            )

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Set application style (optional)
    app.setStyle("Fusion")
    
    window = ExcelReaderWindow()
    window.show()
    sys.exit(app.exec())