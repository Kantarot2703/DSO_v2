from PyQt5 import QtWidgets
from ui.pdf_viewer import PDFViewer
from checklist_loader import load_checklist
from pdf_reader import extract_text_by_page
from checker import check_term_in_page
from result_exporter import export_result_to_excel
import os

class DSOApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("DSO - Digital Sign Out Checker")
        self.setGeometry(100, 100, 1000, 700)
        self.excel_path = ""
        self.pdf_path = ""
        self.checklist_df = None
        self.pages = None
        self.results = []
        self.init_ui()

    def init_ui(self):
        layout = QtWidgets.QVBoxLayout()

        # Upload buttons
        file_layout = QtWidgets.QHBoxLayout()
        self.pdf_btn = QtWidgets.QPushButton("📄 Upload PDF (Artwork)")
        self.excel_btn = QtWidgets.QPushButton("📋 Upload Checklist (Excel)")
        self.pdf_label = QtWidgets.QLabel("PDF: Not selected")
        self.excel_label = QtWidgets.QLabel("Checklist: Not selected")
        self.pdf_btn.clicked.connect(self.load_pdf)
        self.excel_btn.clicked.connect(self.load_excel)
        file_layout.addWidget(self.pdf_btn)
        file_layout.addWidget(self.excel_btn)
        file_layout.addStretch()

        # Search bar
        search_layout = QtWidgets.QHBoxLayout()
        self.search_input = QtWidgets.QLineEdit()
        self.search_input.setPlaceholderText("Search term (Ctrl+F)...")
        self.search_btn = QtWidgets.QPushButton("Search")
        self.search_btn.clicked.connect(self.preview_pdf)
        search_layout.addWidget(self.search_input)
        search_layout.addWidget(self.search_btn)

        # Result Table
        self.result_table = QtWidgets.QTableWidget()
        self.result_table.setColumnCount(6)
        self.result_table.setHorizontalHeaderLabels(["Page", "Language", "Term", "Found", "Passed", "Reason"])
        self.result_table.horizontalHeader().setStretchLastSection(True)

        # Action Buttons
        action_layout = QtWidgets.QHBoxLayout()
        self.check_btn = QtWidgets.QPushButton("Start Checking")
        self.export_btn = QtWidgets.QPushButton("Export to Excel")
        self.preview_btn = QtWidgets.QPushButton("Preview PDF")
        self.check_btn.clicked.connect(self.start_checking)
        self.export_btn.clicked.connect(self.export_results)
        self.preview_btn.clicked.connect(self.preview_pdf)
        action_layout.addWidget(self.check_btn)
        action_layout.addWidget(self.export_btn)
        action_layout.addWidget(self.preview_btn)
        action_layout.addStretch()

        # Add to layout
        layout.addLayout(file_layout)
        layout.addWidget(self.pdf_label)
        layout.addWidget(self.excel_label)
        layout.addLayout(search_layout)
        layout.addWidget(self.result_table)
        layout.addLayout(action_layout)
        self.setLayout(layout)

    def load_pdf(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Select PDF Artwork", "", "PDF Files (*.pdf)")
        if path:
            self.pdf_path = path
            self.pdf_label.setText(f"PDF: {os.path.basename(path)}")
            self.pages = extract_text_by_page(path)

    def load_excel(self):
        if not self.pdf_path:
            QtWidgets.QMessageBox.warning(self, "PDF Required", "Please upload a PDF file first.")
            return

        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Select Checklist Excel", "", "Excel Files (*.xlsx *.xls)")
        if path:
            self.excel_path = path
            self.excel_label.setText(f"Checklist: {os.path.basename(path)}")
            try:
                self.checklist_df = load_checklist(path, os.path.basename(self.pdf_path))
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Checklist Load Error", str(e))

    def start_checking(self):
        if self.checklist_df is None or self.checklist_df.empty or not self.pages:
            QtWidgets.QMessageBox.warning(self, "Missing File", "Please upload both Checklist and PDF before checking.")
            return

        self.results.clear()
        self.result_table.setRowCount(0)

        for _, row in self.checklist_df.iterrows():
            for lang in row['Language List']:
                term = row['Term (Text)']
                for page_num, page_items in enumerate(self.pages, start=1):
                    result = check_term_in_page(term, page_items, row)
                    self.results.append({
                        "Page": page_num,
                        "Language": lang,
                        "Term": term,
                        "Found": result["found"],
                        "Passed": result["matched"],
                        "Reason": "; ".join(result["reasons"]),
                    })

        self.display_results()

    def display_results(self):
        self.result_table.setRowCount(len(self.results))
        for row_idx, res in enumerate(self.results):
            self.result_table.setItem(row_idx, 0, QtWidgets.QTableWidgetItem(str(res["Page"])))
            self.result_table.setItem(row_idx, 1, QtWidgets.QTableWidgetItem(res["Language"]))
            self.result_table.setItem(row_idx, 2, QtWidgets.QTableWidgetItem(res["Term"]))
            self.result_table.setItem(row_idx, 3, QtWidgets.QTableWidgetItem(str(res["Found"])))
            self.result_table.setItem(row_idx, 4, QtWidgets.QTableWidgetItem(str(res["Passed"])))
            self.result_table.setItem(row_idx, 5, QtWidgets.QTableWidgetItem(res["Reason"]))

    def export_results(self):
        if not self.results:
            QtWidgets.QMessageBox.warning(self, "No Results", "Please run checking before exporting.")
            return
        export_result_to_excel(self.results)

    def preview_pdf(self):
        if self.pdf_path:
            term = self.search_input.text().strip()
            if not term:
                QtWidgets.QMessageBox.information(self, "Enter Term", "Please enter a term to highlight in PDF.")
                return
            viewer = PDFViewer(self.pdf_path, search_term=term)
            viewer.show()
        else:
            QtWidgets.QMessageBox.warning(self, "No PDF", "Please upload a PDF first.")
