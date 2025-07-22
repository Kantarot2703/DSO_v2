from PyQt5.QtWidgets import QWidget, QVBoxLayout, QGraphicsScene, QGraphicsView
from PyQt5.QtGui import QPixmap, QImage, QPen, QColor
from PyQt5.QtCore import QRectF
import fitz  # PyMuPDF
import io

class PDFViewer(QWidget):
    def __init__(self, pdf_path, search_term=None):
        super().__init__()
        self.setWindowTitle("PDF Preview with Highlight")
        self.resize(800, 1000)
        self.search_term = search_term
        self.doc = fitz.open(pdf_path)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        self.viewer = QGraphicsView()
        self.scene = QGraphicsScene()
        self.viewer.setScene(self.scene)
        layout.addWidget(self.viewer)
        self.setLayout(layout)

        self.render_pages()

    def render_pages(self):
        for page_number in range(len(self.doc)):
            page = self.doc.load_page(page_number)
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # High-res
            img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGBA8888)
            pixmap = QPixmap.fromImage(img)

            self.scene.addPixmap(pixmap)
            self.scene.setSceneRect(pixmap.rect())

            # Highlight
            if self.search_term:
                highlights = page.search_for(self.search_term)
                pen = QPen(QColor("yellow"))
                pen.setWidth(3)

                for rect in highlights:
                    highlight_rect = QRectF(rect.x0 * 2, rect.y0 * 2, (rect.x1 - rect.x0) * 2, (rect.y1 - rect.y0) * 2)
                    self.scene.addRect(highlight_rect, pen)
