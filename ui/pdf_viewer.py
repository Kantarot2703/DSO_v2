from PyQt5.QtWidgets import QWidget, QVBoxLayout, QGraphicsScene, QGraphicsView, QPushButton, QLabel, QHBoxLayout
from PyQt5.QtGui import QPixmap, QImage, QPen, QColor
from PyQt5.QtCore import QRectF
import fitz  # PyMuPDF
import io

class PDFViewer(QWidget):
    def __init__(self, pdf_path, search_term=None, product_infos=None):
        super().__init__()
        self.setWindowTitle("PDF Preview with Highlight")
        self.resize(800, 1000)

        self.pdf_path = pdf_path
        self.search_term = search_term
        self.product_infos = product_infos or []
        self.doc = fitz.open(pdf_path)
        self.current_page = 0

        self.init_ui()
        self.render_page()

    def init_ui(self):
        layout = QVBoxLayout()

        # Product info label
        self.product_label = QLabel("🧾 Product: -\n🔢 Part No: -\n🔁 Rev: -")
        layout.addWidget(self.product_label)

        # PDF display
        self.viewer = QGraphicsView()
        self.scene = QGraphicsScene()
        self.viewer.setScene(self.scene)
        layout.addWidget(self.viewer)

        # Navigation buttons
        nav_layout = QHBoxLayout()
        self.prev_btn = QPushButton("⬅ Previous")
        self.next_btn = QPushButton("Next ➡")
        self.prev_btn.clicked.connect(self.show_prev_page)
        self.next_btn.clicked.connect(self.show_next_page)
        nav_layout.addWidget(self.prev_btn)
        nav_layout.addWidget(self.next_btn)
        layout.addLayout(nav_layout)

        self.setLayout(layout)

    def render_page(self):
        self.scene.clear()

        page = self.doc.load_page(self.current_page)
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # High resolution
        img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGBA8888)
        pixmap = QPixmap.fromImage(img)
        self.scene.addPixmap(pixmap)
        self.scene.setSceneRect(pixmap.rect())

        # Set product info label
        info = next((p for p in self.product_infos if p["page"] == self.current_page + 1), {})
        self.product_label.setText(
            f"🧾 Product: {info.get('product_name', '-')}\n"
            f"🔢 Part No: {info.get('part_no', '-')}\n"
            f"🔁 Rev: {info.get('rev', '-')}"
        )

        # Highlight search term
        if self.search_term:
            highlights = page.search_for(self.search_term)
            pen = QPen(QColor("yellow"))
            pen.setWidth(3)
            for rect in highlights:
                highlight_rect = QRectF(rect.x0 * 2, rect.y0 * 2, (rect.x1 - rect.x0) * 2, (rect.y1 - rect.y0) * 2)
                self.scene.addRect(highlight_rect, pen)

    def show_next_page(self):
        if self.current_page < len(self.doc) - 1:
            self.current_page += 1
            self.render_page()

    def show_prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.render_page()
