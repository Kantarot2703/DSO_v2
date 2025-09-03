import re
import fitz
from typing import List, Dict, Optional
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtCore import Qt, QRectF
from PyQt5.QtGui import QPixmap, QPen, QColor, QBrush
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGraphicsScene, QGraphicsView,
    QLabel, QPushButton, QSizePolicy
)

class PDFViewer(QWidget):
    def __init__(
        self,
        pdf_path: str,
        search_term: Optional[str] = None,
        product_infos: Optional[List[Dict]] = None,
        per_word_highlight: bool = True,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)

        self.pdf_path = pdf_path
        self.search_term = (search_term or "").strip()
        self.product_infos = product_infos or []
        self.per_word_highlight = per_word_highlight

        self.doc = fitz.open(self.pdf_path)
        self.page_count = self.doc.page_count
        self.current_page = 0 
        self.zoom = 1.75     

        self.setWindowTitle("PDF Preview with Highlight")
        self.resize(1000, 800)
        self.init_ui()
        self.render_page()

    # UI 
    def init_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)

        self.info_label = QLabel("")
        self.info_label.setWordWrap(True)
        root.addWidget(self.info_label)

        # แถบเครื่องมือ
        tools = QHBoxLayout()
        self.prev_btn = QPushButton("◀ Prev")
        self.next_btn = QPushButton("Next ▶")
        self.zoom_out_btn = QPushButton("–")
        self.zoom_in_btn = QPushButton("+")
        self.fit_btn = QPushButton("Fit width")
        self.page_label = QLabel("")

        self.prev_btn.clicked.connect(self.go_prev)
        self.next_btn.clicked.connect(self.go_next)
        self.zoom_out_btn.clicked.connect(lambda: self.set_zoom(self.zoom * 0.9))
        self.zoom_in_btn.clicked.connect(lambda: self.set_zoom(self.zoom * 1.1))
        self.fit_btn.clicked.connect(self.fit_width)

        for w in (self.prev_btn, self.next_btn, self.zoom_out_btn, self.zoom_in_btn, self.fit_btn):
            w.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
            tools.addWidget(w)
        tools.addStretch(1)
        tools.addWidget(self.page_label)
        root.addLayout(tools)

        # พื้นที่แสดงภาพ PDF
        self.scene = QGraphicsScene(self)
        self.view = QGraphicsView(self.scene)
        self.view.setRenderHints(self.view.renderHints())
        self.view.setDragMode(QGraphicsView.ScrollHandDrag)
        self.view.setViewportUpdateMode(QGraphicsView.SmartViewportUpdate)
        self.view.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        root.addWidget(self.view, 1)

    # Navigation
    def update_nav_state(self):
        self.page_label.setText(f"Page {self.current_page + 1} / {self.page_count}")
        self.prev_btn.setEnabled(self.current_page > 0)
        self.next_btn.setEnabled(self.current_page < self.page_count - 1)

        info = next((p for p in self.product_infos if int(p.get("page", -1)) == self.current_page + 1), {})
        prod = info.get("product_name", "-")
        part = info.get("part_no", "-")
        rev  = info.get("rev", "-")
        self.info_label.setText(f"🧾 Product: {prod}   |   🔢 Part No: {part}   |   🔁 Rev: {rev}")

    def go_prev(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.render_page()

    def go_next(self):
        if self.current_page < self.page_count - 1:
            self.current_page += 1
            self.render_page()

    def set_zoom(self, z: float):
        self.zoom = max(0.2, min(6.0, z))
        self.render_page()

    def fit_width(self):

        page = self.doc.load_page(self.current_page)
        pm = page.get_pixmap(matrix=fitz.Matrix(1, 1), alpha=False)
        if pm.width == 0:
            return
        
        viewport_w = max(50, self.view.viewport().width() - 16)
        self.zoom = max(0.2, min(6.0, viewport_w / pm.width))
        self.render_page()

    # Render
    def render_page(self):
        self.scene.clear()
        self.update_nav_state()

        page = self.doc.load_page(self.current_page)
        matrix = fitz.Matrix(self.zoom, self.zoom)

        pm = page.get_pixmap(matrix=matrix, alpha=False)
        png_bytes = pm.tobytes("png")
        pixmap = QPixmap()
        pixmap.loadFromData(png_bytes)

        self.scene.addPixmap(pixmap)
        self.scene.setSceneRect(QRectF(pixmap.rect()))

        if self.search_term:
            rects = []

            def _search_boxes(p, text):
                try:
                    return p.search_for(text)
                except Exception:
                    return []

            rects.extend(_search_boxes(page, self.search_term))

            if self.per_word_highlight:
                tokens = re.findall(r"[A-Za-z0-9\u00C0-\u017F\+\-/]+", self.search_term)
                for tok in tokens:
                    if len(tok) >= 2:
                        rects.extend(_search_boxes(page, tok))

            pen = QPen(Qt.NoPen)
            brush = QBrush(QColor(255, 235, 59, 120)) 

            # วาดทับข้อความ
            for r in rects:
                rr = QRectF(r.x0 * self.zoom, r.y0 * self.zoom,
                            (r.x1 - r.x0) * self.zoom, (r.y1 - r.y0) * self.zoom)
                self.scene.addRect(rr, pen, brush)

        # scroll กลับขึ้นมุมซ้ายบนทุกครั้งที่เปลี่ยนหน้า/ซูม
        self.view.centerOn(0, 0)

    def closeEvent(self, e):
        try:
            self.doc.close()
        except Exception:
            pass
        super().closeEvent(e)
