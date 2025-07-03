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
        
        # Document Generator Tabs
        self.setup_document_tab('dati_generali_procedura', "Genera Documenti Dati Generali")
        self.setup_document_tab('generazioni_offerte', "Genera Documenti Offerte")
        
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
        excel_layout.setContentsMargins(10, 10, 10, 10)
        excel_layout.setSpacing(10)
        excel_tab.setLayout(excel_layout)
        
        # File selection
        file_group = QGroupBox("Selezione File Excel")
        file_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                padding: 5px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px;
            }
        """)
        file_layout = QHBoxLayout()
        file_layout.setContentsMargins(5, 5, 5, 5)
        self.file_path = QLineEdit()
        self.file_path.setPlaceholderText("Seleziona file Excel...")
        file_button = QPushButton("Sfoglia...")
        file_button.setIcon(QIcon.fromTheme("document-open"))
        file_button.clicked.connect(self.browse_excel_file)
        file_layout.addWidget(self.file_path)
        file_layout.addWidget(file_button)
        file_group.setLayout(file_layout)
        
        # Splitter for the two lists
        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        self.splitter.setHandleWidth(8)
        self.splitter.setStyleSheet("""
            QSplitter::handle {
                background: #dee2e6;
                width: 4px;
                margin: 2px;
            }
        """)
        
        # List for dati_generali_procedura
        self.dati_generali_list = QListWidget()
        self.dati_generali_list.itemDoubleClicked.connect(
            lambda item: self.show_variable_value(item, 'dati_generali_procedura'))
        self.dati_generali_list.setMinimumWidth(300)
        
        # List for generazioni_offerte
        self.generazioni_offerte_list = QListWidget()
        self.generazioni_offerte_list.itemDoubleClicked.connect(
            lambda item: self.show_variable_value(item, 'generazioni_offerte'))
        self.generazioni_offerte_list.setMinimumWidth(300)
        
        # Add lists to splitter
        self.splitter.addWidget(self.create_list_group(self.dati_generali_list, "Dati Generali Procedura"))
        self.splitter.addWidget(self.create_list_group(self.generazioni_offerte_list, "Generazioni Offerte"))
        
        # Set initial sizes and stretch factors
        self.splitter.setSizes([400, 400])
        self.splitter.setStretchFactor(0, 1)
        self.splitter.setStretchFactor(1, 1)
        
        # Results display
        self.results_display = QTextEdit()
        self.results_display.setReadOnly(True)
        self.results_display.setStyleSheet("""
            QTextEdit {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                padding: 8px;
                min-height: 120px;
                font-size: 12px;
            }
        """)
        
        # Add widgets to Excel tab
        excel_layout.addWidget(file_group)
        excel_layout.addWidget(self.splitter)
        excel_layout.addWidget(QLabel("Output:"))
        excel_layout.addWidget(self.results_display)
        
        # Scan button
        scan_button = QPushButton("Leggi Fogli Excel")
        scan_button.setIcon(QIcon.fromTheme("document-open"))
        scan_button.setStyleSheet("""
            QPushButton {
                padding: 8px;
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
                margin: 5px;
                padding: 5px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px;
            }
        """)
        layout = QVBoxLayout()
        layout.setContentsMargins(5, 5, 5, 5)
        
        list_widget.setStyleSheet("""
            QListWidget {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                padding: 5px;
                font-size: 12px;
            }
            QListWidget::item {
                padding: 3px;
            }
        """)
        
        layout.addWidget(list_widget)
        group.setLayout(layout)
        return group

    def browse_excel_file(self):
        """Open a file dialog to select an Excel file"""
        file_name, _ = QFileDialog.getOpenFileName(
            self, 
            "Apri file Excel", 
            "", 
            "Excel Files (*.xlsx *.xls)"
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
                    
                    # Special processing for dati_generali_procedura sheet
                    if sheet_name == 'dati_generali_procedura':
                        # Process column C with names from column E, starting from row 2
                        for row_idx, row in enumerate(sheet.iter_rows(min_row=2, min_col=3, max_col=5, values_only=False), start=2):
                            cell_c = row[0]  # Column C
                            cell_e = row[2]  # Column E
                            
                            # Skip if column C or E is empty or contains only whitespace
                            if (cell_c.value is None or not str(cell_c.value).strip() or 
                                cell_e.value is None or not str(cell_e.value).strip()):
                                continue
                                
                            # Get name from column E
                            name = str(cell_e.value)
                            item_text = f"{name}: {cell_c.value}"
                            
                            item = QListWidgetItem(item_text)
                            item.setData(Qt.ItemDataRole.UserRole, {
                                'type': 'value',
                                'coord': cell_c.coordinate,
                                'name': name,
                                'value': cell_c.value
                            })
                            target_list.addItem(item)
                        
                        # Process column D with flag names, starting from row 2
                        for row_idx, row in enumerate(sheet.iter_rows(min_row=2, min_col=4, max_col=4, values_only=False), start=2):
                            cell_d = row[0]  # Column D
                            
                            if cell_d.value is not None and str(cell_d.value).strip():
                                flag_name = f"flag_{row_idx}"
                                item_text = f"{flag_name}: {cell_d.value}"
                                
                                item = QListWidgetItem(item_text)
                                item.setData(Qt.ItemDataRole.UserRole, {
                                    'type': 'flag',
                                    'coord': cell_d.coordinate,
                                    'name': flag_name,
                                    'value': cell_d.value
                                })
                                target_list.addItem(item)
                    
                    else:
                        # Normal processing for other sheets
                        for row in sheet.iter_rows():
                            for cell in row:
                                if cell.value is not None and isinstance(cell.value, str) and cell.value.strip():
                                    item = QListWidgetItem(f"{cell.coordinate}: {cell.value}")
                                    item.setData(Qt.ItemDataRole.UserRole, {
                                        'type': 'standard',
                                        'coord': cell.coordinate,
                                        'value': cell.value
                                    })
                                    target_list.addItem(item)
                        
                        # Pre-fill QLineEdit fields from generazioni_offerte sheet
                        self.prefill_from_generazioni_offerte(sheet)
                
                    self.results_display.append(f"Trovate {target_list.count()} variabili nel foglio {sheet_name}")
                
                except KeyError:
                    QMessageBox.warning(self, "Attenzione", f"Il foglio '{sheet_name}' non esiste nel file.")
                    continue
            
            workbook.close()
            
        except Exception as e:
            QMessageBox.critical(self, "Errore", f"Errore nella lettura del file Excel:\n{str(e)}")

    def prefill_from_generazioni_offerte(self, sheet):
        """Prefill QLineEdit fields from generazioni_offerte sheet"""
        # Define mapping between cell coordinates and field names
        field_mapping = {
            'B2': 'oggetto_fornitura_servizio',
            'B3': 'nome_ditta',
            'B4': 'indirizzo_ditta',
            'B5': 'cap_ditta',
            'B6': 'pec_ditta'
        }
        
        for coord, field_name in field_mapping.items():
            try:
                cell = sheet[coord]
                if cell.value is not None:
                    self.doc_fields[field_name].setText(str(cell.value))
            except:
                continue

    def show_variable_value(self, item, sheet_name):
        if not self.current_file or sheet_name not in self.sheet_data or not self.sheet_data[sheet_name]:
            return
            
        item_data = item.data(Qt.ItemDataRole.UserRole)
        
        try:
            sheet = self.sheet_data[sheet_name]
            
            if item_data['type'] == 'value':
                self.results_display.append(
                    f"\nDettaglio variabile (valore):\n"
                    f"• Nome: {item_data['name']}\n"
                    f"• Posizione: {item_data['coord']}\n"
                    f"• Valore: {item_data['value']}\n"
                    f"• Tipo: {type(item_data['value']).__name__}\n"
                    f"----------------------------")
            
            elif item_data['type'] == 'flag':
                self.results_display.append(
                    f"\nDettaglio flag:\n"
                    f"• Nome: {item_data['name']}\n"
                    f"• Posizione: {item_data['coord']}\n"
                    f"• Valore: {item_data['value']}\n"
                    f"• Tipo: {type(item_data['value']).__name__}\n"
                    f"----------------------------")
            else:
                # Normal processing for other sheets
                cell_ref = item_data['coord']
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

    def setup_document_tab(self, sheet_type, tab_name):
        """Setup a Document Generator tab for specific sheet type"""
        doc_tab = QWidget()
        doc_layout = QVBoxLayout()
        doc_layout.setContentsMargins(10, 10, 10, 10)
        doc_layout.setSpacing(10)
        doc_tab.setLayout(doc_layout)
        
        # Template selection
        template_group = QGroupBox("Configurazione Documento")
        template_group.setStyleSheet("QGroupBox { font-weight: bold; }")
        template_layout = QFormLayout()
        
        self.template_path = QLineEdit()
        self.template_path.setPlaceholderText("Seleziona template Word (multipli)...")
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
        
        # Document fields (common fields for both tabs)
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
        
        # Add only relevant fields based on sheet type
        if sheet_type == 'dati_generali_procedura':
            fields_to_show = ['numero_CUP', 'servizio_fornitura', 'prestazione_servizio_fornitura', 
                             'nome_cognome', 'mail_contatto', 'acronimo_progetto']
        else:
            fields_to_show = ['oggetto_fornitura_servizio', 'nome_ditta', 'indirizzo_ditta', 
                            'cap_ditta', 'pec_ditta']
        
        for field_name in fields_to_show:
            label = field_name.replace('_', ' ').title() + ":"
            self.doc_fields[field_name].setPlaceholderText(f"Inserisci {label.lower()}")
            template_layout.addRow(label, self.doc_fields[field_name])
        
        template_group.setLayout(template_layout)
        
        # Generate button
        generate_btn = QPushButton(f"Genera Documento {tab_name.split()[-1]}")
        generate_btn.setIcon(QIcon.fromTheme("document-save-as"))
        generate_btn.setStyleSheet("""
            QPushButton {
                padding: 8px;
                background-color: #2980b9;
                color: white;
                border-radius: 10px;
            }
            QPushButton:hover {
                background-color: #3498db;
            }
        """)
        generate_btn.clicked.connect(lambda: self.generate_document(sheet_type))
        
        # Add widgets to Document tab
        doc_layout.addWidget(template_group)
        doc_layout.addWidget(generate_btn)
        doc_layout.addStretch()
        
        self.tabs.addTab(doc_tab, tab_name)

    def browse_template_file(self):
        """Open a file dialog to select multiple Word templates"""
        file_names, _ = QFileDialog.getOpenFileNames(
            self, 
            "Seleziona template Word", 
            "", 
            "Word Documents (*.docx)"
        )
        if file_names:
            self.template_path.setText("; ".join(file_names))

    def browse_output_dir(self):
        dir_name = QFileDialog.getExistingDirectory(
            self, "Seleziona cartella di output"
        )
        if dir_name:
            self.output_dir.setText(dir_name)

    def generate_document(self, source_sheet):
        """Generate Word documents from multiple templates using data from specified sheet"""
        template_paths = [path.strip() for path in self.template_path.text().split(";") if path.strip()]
        output_dir = self.output_dir.text()
        
        if not template_paths:
            QMessageBox.warning(self, "Attenzione", "Seleziona almeno un template Word.")
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
            if source_sheet == 'dati_generali_procedura':
                # Add values and flags from the list
                for i in range(self.dati_generali_list.count()):
                    item = self.dati_generali_list.item(i)
                    item_data = item.data(Qt.ItemDataRole.UserRole)
                    context[item_data['name']] = item_data['value']
            else:
                # Add data from generazioni_offerte sheet
                for i in range(self.generazioni_offerte_list.count()):
                    item = self.generazioni_offerte_list.item(i)
                    item_data = item.data(Qt.ItemDataRole.UserRole)
                    if isinstance(item_data, dict):
                        context[item_data['name']] = item_data['value']
                    else:
                        sheet = self.sheet_data[source_sheet]
                        cell = sheet[item_data]
                        context[f"{source_sheet}_{item_data}"] = cell.value
            
            # Generate documents for each template
            generated_files = []
            for template_path in template_paths:
                try:
                    # Generate filename
                    cognome = context['nome_cognome'].split()[-1] if context['nome_cognome'] else 'documento'
                    template_name = os.path.splitext(os.path.basename(template_path))[0]
                    source_suffix = "DatiGenerali" if source_sheet == 'dati_generali_procedura' else "Offerte"
                    output_filename = f"RichiestaOfferta_{cognome}_{source_suffix}_{template_name}.docx"
                    output_path = os.path.join(output_dir, output_filename)
                    
                    # Render and save document
                    doc = DocxTemplate(template_path)
                    doc.render(context)
                    doc.save(output_path)
                    generated_files.append(output_path)
                    
                except Exception as e:
                    QMessageBox.warning(
                        self, 
                        "Attenzione", 
                        f"Errore durante la generazione del documento {template_path}:\n{str(e)}"
                    )
                    continue
            
            if generated_files:
                success_message = "Documenti generati con successo:\n\n" + "\n".join(generated_files)
                QMessageBox.information(self, "Successo", success_message)
            else:
                QMessageBox.warning(self, "Attenzione", "Nessun documento è stato generato.")
                
        except Exception as e:
            QMessageBox.critical(
                self, 
                "Errore", 
                f"Errore durante la generazione dei documenti:\n{str(e)}"
            )

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = ExcelReaderWindow()
    window.show()
    sys.exit(app.exec())