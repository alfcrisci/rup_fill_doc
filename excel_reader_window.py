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
        self.setGeometry(100, 100, 1300, 800)
        
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
        self.generazioni_offerte_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        self.generazioni_offerte_list.itemDoubleClicked.connect(
            lambda item: self.show_variable_value(item, 'generazioni_offerte'))
        self.generazioni_offerte_list.setMinimumWidth(300)
        
        # Add lists to splitter
        self.splitter.addWidget(self.create_list_group(self.dati_generali_list, "Dati Procedura"))
        self.splitter.addWidget(self.create_list_group(self.generazioni_offerte_list, "Generazioni Richeste Offerte - Selezione multipla con click"))
        
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
            QListWidget::item:selected {
                background-color: #d4e6f1;
                color: #2c3e50;
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
            
            # Read dati_generali_procedura sheet (manteniamo la logica originale)
            if 'dati_generali_procedura' in workbook.sheetnames:
                sheet = workbook['dati_generali_procedura']
                self.sheet_data['dati_generali_procedura'] = sheet
                
                # Mappatura tra nomi delle colonne Excel e campi QLineEdit
                field_mapping = {
                    'numero_CUP': None,
                    'servizio_fornitura': None,
                    'acronimo_progetto': None,
                    'oggetto_fornitura_servizio': None,
                    'oggetto_esteso_fornitura_servizio': None,
                    'nome_cognome_richiedente': None,
                    'mail_contatto_richiedente': None
                }
                
                # Process column C with names from column E, starting from row 2
                for row_idx, row in enumerate(sheet.iter_rows(min_row=2, min_col=3, max_col=5, values_only=False), start=2):
                    cell_c = row[0]  # Column C
                    cell_e = row[2]  # Column E
                    
                    if (cell_c.value is None or not str(cell_c.value).strip() or 
                        cell_e.value is None or not str(cell_e.value).strip()):
                        continue
                        
                    name = str(cell_e.value)
                    item_text = f"{name}: {cell_c.value}"
                    
                    # Controlla se questo valore corrisponde a uno dei nostri campi
                    for field_name in field_mapping.keys():
                        if field_name.lower() in name.lower():
                            field_mapping[field_name] = str(cell_c.value) if cell_c.value else ""
                    
                    item = QListWidgetItem(item_text)
                    item.setData(Qt.ItemDataRole.UserRole, {
                        'type': 'value',
                        'coord': cell_c.coordinate,
                        'name': name,
                        'value': cell_c.value
                    })
                    self.dati_generali_list.addItem(item)
                
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
                        self.dati_generali_list.addItem(item)
                
                # Riempimento automatico dei campi QLineEdit
                for field_name, value in field_mapping.items():
                    if value is not None and field_name in self.doc_fields:
                        self.doc_fields[field_name].setText(value)
            
            # Read generazioni_offerte sheet come tabella (prima riga = nomi colonne)
            if 'generazioni_offerte' in workbook.sheetnames:
                sheet = workbook['generazioni_offerte']
                self.sheet_data['generazioni_offerte'] = sheet
                
                # Leggi la prima riga come nomi delle colonne
                headers = []
                for cell in sheet[1]:
                    header_name = str(cell.value) if cell.value else f"col_{cell.column_letter}"
                    headers.append(header_name)
                
                # Leggi i dati per ogni riga successiva
                for row_idx, row in enumerate(sheet.iter_rows(min_row=2), start=2):
                    row_data = {}
                    for col_idx, cell in enumerate(row, start=1):
                        if col_idx - 1 < len(headers):
                            row_data[headers[col_idx - 1]] = cell.value
                    
                    # Aggiungi alla lista come voce unica per la riga
                    item_text = f" {', '.join(f' {v}' for k, v in row_data.items() if v)}"
                    item = QListWidgetItem(item_text)
                    item.setData(Qt.ItemDataRole.UserRole, {
                        'type': 'row',
                        'row_idx': row_idx-1,
                        'data': row_data
                    })
                    self.generazioni_offerte_list.addItem(item)
            
            workbook.close()
            
        except Exception as e:
            QMessageBox.critical(self, "Errore", f"Errore nella lettura del file Excel:\n{str(e)}")

    def show_variable_value(self, item, sheet_name):
        if not self.current_file or sheet_name not in self.sheet_data or not self.sheet_data[sheet_name]:
            return
            
        item_data = item.data(Qt.ItemDataRole.UserRole)
        
        try:
            if sheet_name == 'generazioni_offerte' and item_data['type'] == 'row':
                # Mostra tutti i dati della riga
                row_data = item_data['data']
                details = "\n".join([f"• {k}: {v} ({type(v).__name__})" for k, v in row_data.items()])
                self.results_display.append(
                    f"\nDettaglio riga {item_data['row_idx']}:\n"
                    f"{details}\n"
                    f"----------------------------")
            else:
                # Mostra dettagli per dati_generali_procedura o celle singole
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
            'nome_cognome_richiedente': QLineEdit(),
            'mail_contatto_richiedente': QLineEdit(),
            'acronimo_progetto': QLineEdit(),
            'oggetto_fornitura_servizio': QLineEdit(),
            'oggetto_esteso_fornitura_servizio': QLineEdit(),
            'nome_ditta': QLineEdit(),
            'indirizzo_ditta': QLineEdit(),
            'cap_ditta': QLineEdit(),
            'pec_ditta': QLineEdit()
        }
        
        # Add only relevant fields based on sheet type
        if sheet_type == 'dati_generali_procedura':
            fields_to_show = [ 
                'servizio_fornitura', 
                'acronimo_progetto',
                'numero_CUP',
                'oggetto_fornitura_servizio',
                'oggetto_esteso_fornitura_servizio',                                
                'nome_cognome_richiedente', 
                'mail_contatto_richiedente'
            ]
        else:
            fields_to_show = [ 
                'servizio_fornitura', 
                'acronimo_progetto',
                'numero_CUP',
                'oggetto_fornitura_servizio',
                'oggetto_esteso_fornitura_servizio',                                
                'nome_cognome_richiedente', 
                'mail_contatto_richiedente'
            ]
        
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
            
            # Prepare base context (dati_generali + form fields)
            context = {
                'numero_CUP': self.doc_fields['numero_CUP'].text(),
                'servizio_fornitura': self.doc_fields['servizio_fornitura'].text(),
                'prestazione_servizio_fornitura': self.doc_fields['prestazione_servizio_fornitura'].text(),
                'nome_cognome_richiedente': self.doc_fields['nome_cognome_richiedente'].text(),
                'mail_contatto_richiedente': self.doc_fields['mail_contatto_richiedente'].text(),
                'acronimo_progetto': self.doc_fields['acronimo_progetto'].text(),
                'oggetto_fornitura_servizio': self.doc_fields['oggetto_fornitura_servizio'].text(),
                'oggetto_esteso_fornitura_servizio': self.doc_fields['oggetto_esteso_fornitura_servizio'].text(),
                'nome_ditta': self.doc_fields['nome_ditta'].text(),
                'indirizzo_ditta': self.doc_fields['indirizzo_ditta'].text(),
                'cap_ditta': self.doc_fields['cap_ditta'].text(),
                'pec_ditta': self.doc_fields['pec_ditta'].text(),
                'data_corrente': datetime.now().strftime('%d/%m/%Y')
            }
            
            # Add data from dati_generali_procedura sheet
            for i in range(self.dati_generali_list.count()):
                item = self.dati_generali_list.item(i)
                item_data = item.data(Qt.ItemDataRole.UserRole)
                context[item_data['name']] = item_data['value']
            
            # Special processing for generazioni_offerte sheet
            if source_sheet == 'generazioni_offerte':
                # Genera un documento per ogni riga selezionata
                selected_items = self.generazioni_offerte_list.selectedItems()
                if not selected_items:
                    QMessageBox.warning(self, "Attenzione", "Seleziona almeno una riga dal foglio generazioni_offerte.")
                    return
                
                generated_files = []
                for item in selected_items:
                    item_data = item.data(Qt.ItemDataRole.UserRole)
                    row_context = context.copy()
                    row_context.update(item_data['data'])
                    
                    # Generate documents for each template
                    for template_path in template_paths:
                        try:
                            # Generate filename
                            cognome = row_context['nome_cognome_richiedente'].split()[-1] if row_context['nome_cognome_richiedente'] else 'documento'
                            progetto = row_context['acronimo_progetto']
                            template_name = os.path.splitext(os.path.basename(template_path))[0]
                            output_filename = f"Richiesta_Offerta_{cognome}_{progetto}_ditta_{item_data['row_idx']}.docx"
                            output_path = os.path.join(output_dir, output_filename)
                            
                            # Render and save document
                            doc = DocxTemplate(template_path)
                            doc.render(row_context)
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
            
            else:
                # Original processing for dati_generali_procedura
                generated_files = []
                for template_path in template_paths:
                    try:
                        # Generate filename
                        cognome = context['nome_cognome_richiedente'].split()[-1] if context['nome_cognome_richiedente'] else 'documento'
                        progetto = context['acronimo_progetto']
                        template_name = os.path.splitext(os.path.basename(template_path))[0]
                        output_filename = f"{template_name}_{cognome}_{progetto}.docx"
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