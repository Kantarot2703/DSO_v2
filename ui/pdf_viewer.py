from typing import List, Dict, Optional
from collections import OrderedDict
import re
import fitz 
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtCore import Qt, QRectF, QTimer
from PyQt5.QtGui import QPixmap, QPen, QColor, QBrush, QKeySequence
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGraphicsScene, QGraphicsView,
    QLabel, QPushButton, QSizePolicy, QListWidget, QListWidgetItem, QComboBox,
    QMainWindow, QShortcut, QSplitter, QScrollBar
)

# ------------------------------ Configs ------------------------------
ARTWORK_MARGIN_L = 0.08
ARTWORK_MARGIN_R = 0.08
ARTWORK_MARGIN_T = 0.12
ARTWORK_MARGIN_B = 0.12

# ความโปร่งของสีไฮไลท์ 0-255
HIGHLIGHT_ALPHA = 80

# แคชผลค้นหาในแต่ละหน้า
DEFAULT_LRU_CAPACITY = 8

# ------------------------------ Helpers ------------------------------
def parse_pages_spec(spec: str, total_pages: int) -> Optional[set]:
    if not spec:
        return None
    s = spec.strip().lower()
    if s in {"-", "all", "all pages", "all page"}:
        return None
    out = set()
    for part in re.split(r"[,\s]+", s):
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            try:
                start = max(1, int(a))
                end = min(total_pages, int(b))
                for p in range(start, end + 1):
                    out.add(p - 1)
            except Exception:
                continue
        else:
            try:
                p = int(part)
                if 1 <= p <= total_pages:
                    out.add(p - 1)
            except Exception:
                continue
    return out or None

def build_terms_from_symbol(symbol_text: str) -> List[str]:
    s = (symbol_text or "").strip()
    if not s or s == "-":
        return []
    terms = [s]
    if "3+" in s or s == "3+":
        for alt in ("3 +", "3＋"):
            if alt not in terms:
                terms.append(alt)
    return terms

def shrink_rect(rect: fitz.Rect,
                left_ratio: float, right_ratio: float,
                top_ratio: float, bottom_ratio: float) -> fitz.Rect:
    w = rect.width
    h = rect.height
    return fitz.Rect(
        rect.x0 + w * left_ratio,
        rect.y0 + h * top_ratio,
        rect.x1 - w * right_ratio,
        rect.y1 - h * bottom_ratio
    )

def rect_center_inside(target: fitz.Rect, container: fitz.Rect) -> bool:
    cx = (target.x0 + target.x1) * 0.5
    cy = (target.y0 + target.y1) * 0.5
    return (container.x0 <= cx <= container.x1) and (container.y0 <= cy <= container.y1)

# ------------------------------ Main Viewer ------------------------------
class _ZoomableGraphicsView(QGraphicsView):
    def __init__(self, outer_viewer, *a, **kw):
        super().__init__(*a, **kw)
        self.outer_viewer = outer_viewer
        self.setDragMode(QGraphicsView.ScrollHandDrag)
        self.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.AnchorUnderMouse)

class PDFViewer(QWidget):
    def __init__(self,
                 pdf_path: str,
                 rows: List[Dict],
                 parent: Optional[QWidget] = None,
                 lru_capacity: int = DEFAULT_LRU_CAPACITY):
        super().__init__(parent)
        self.pdf_path = pdf_path
        self.doc = fitz.open(self.pdf_path)
        self.page_count = len(self.doc)
        self.current_page = 0
        self.zoom = 1.75 

        self.rows = rows or []
        self.filter_mode = "all"
        self.selected_row_id: Optional[int] = None

        self._cache: "OrderedDict[int, Dict[str, List[fitz.Rect]]]" = OrderedDict()
        self._lru_capacity = max(2, int(lru_capacity))
        self._render_busy = False
        self._render_again = False

        self.setWindowTitle("PDF Preview (Requirements evidence)")
        self.resize(1200, 840)
        self._init_ui()
        self._refresh_sidebar()

        QTimer.singleShot(0, self.fit_width)
        QTimer.singleShot(0, self.render_page)
        QTimer.singleShot(0, self._prefetch_neighbors)

    # -------------------------- UI & Sidebar --------------------------
    def _apply_button_style(self, btn: QPushButton):
        btn.setCursor(Qt.PointingHandCursor)
        btn.setStyleSheet(
            """
            QPushButton{
                border: 1px solid #d0d0d0;
                border-radius: 10px;
                padding: 6px 14px;
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #ffffff, stop:1 #f5f5f5);
            }
            QPushButton:hover{ background:#f9f9f9; }
            QPushButton:pressed{ background:#ececec; }
            """
        )
        shadow = QtWidgets.QGraphicsDropShadowEffect(self)
        shadow.setOffset(0, 1)
        shadow.setBlurRadius(10)
        shadow.setColor(QColor(0, 0, 0, 70))
        btn.setGraphicsEffect(shadow)

def _init_ui(self):
    root = QVBoxLayout(self)
    root.setContentsMargins(8, 8, 8, 8)

    # ---- Toolbar ----
    tools = QHBoxLayout()
    self.prev_page_btn = QPushButton("◀ Back")
    self.next_page_btn = QPushButton("Next ▶")
    self.zoom_out_btn  = QPushButton("–")
    self.zoom_in_btn   = QPushButton("+")
    self.fit_btn       = QPushButton("Fit width")
    self.page_label    = QLabel("")
    self.page_label.setStyleSheet("QLabel{font-weight:600; padding-left:8px;}")

    for b in (self.prev_page_btn, self.next_page_btn, self.zoom_out_btn, self.zoom_in_btn, self.fit_btn):
        try:
            self._apply_button_style(b)
        except Exception:
            pass

    # events ของปุ่ม
    self.prev_page_btn.clicked.connect(self.go_prev_page)
    self.next_page_btn.clicked.connect(self.go_next_page)
    self.zoom_out_btn.clicked.connect(lambda: self.set_zoom(self.zoom * 0.9))
    self.zoom_in_btn.clicked.connect(lambda: self.set_zoom(self.zoom * 1.1))
    self.fit_btn.clicked.connect(self.fit_width)

    for w in (self.prev_page_btn, self.next_page_btn, self.zoom_out_btn, self.zoom_in_btn, self.fit_btn, self.page_label):
        w.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        tools.addWidget(w)
    tools.addStretch(1)
    root.addLayout(tools)

    self.splitter = QSplitter(Qt.Vertical)
    self._syncing_pagebar = False

    # Top Viewer area
    top_wrap = QWidget()
    top_v = QVBoxLayout(top_wrap)
    top_v.setContentsMargins(0, 8, 0, 8)

    self.scene = QGraphicsScene(self)
    self.view  = _ZoomableGraphicsView(self) if hasattr(self, '_ZoomableGraphicsView__marker__') else QGraphicsView(self.scene)
    self.view.setScene(self.scene)
    self.view.setRenderHints(QtGui.QPainter.Antialiasing | QtGui.QPainter.SmoothPixmapTransform)
    self.view.setViewportUpdateMode(QGraphicsView.SmartViewportUpdate)
    self.view.setAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
    self.view.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)

    try:
        self.view.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
    except Exception:
        pass

    viewer_row = QHBoxLayout()
    viewer_row.setContentsMargins(0, 0, 0, 0)
    viewer_row.addWidget(self.view, 1)

    # --- Page Scrollbar (เปลี่ยนหน้า) ---
    self.page_bar = QScrollBar(Qt.Vertical)
    self.page_bar.setRange(1, max(1, self.page_count))
    self.page_bar.setPageStep(1)
    self.page_bar.setSingleStep(1)
    self.page_bar.setValue(self.current_page + 1)
    self.page_bar.setFixedWidth(20)
    self.page_bar.valueChanged.connect(self._on_pagebar_changed)

    viewer_row.addWidget(self.page_bar)
    top_v.addLayout(viewer_row, 1)

    self.splitter.addWidget(top_wrap)

    # Bottom: Requirement panel (filter + list)
    bottom_wrap = QWidget()
    bottom_v = QVBoxLayout(bottom_wrap)
    bottom_v.setContentsMargins(0, 0, 0, 0)

    # แถว filter
    top_bar = QHBoxLayout()
    self.filter_combo = QComboBox()
    self.filter_combo.addItems(["All", "Found", "Missing", "Manual"])
    self.filter_combo.currentIndexChanged.connect(self._on_filter_changed)

    self.info_label = QLabel("")

    top_bar.addWidget(QLabel("Filter:"))
    top_bar.addWidget(self.filter_combo)
    top_bar.addStretch(1)
    top_bar.addWidget(self.info_label)
    bottom_v.addLayout(top_bar)

    # รายการ requirement (บรรทัดเดียว/ไม่ตัดบรรทัด)
    self.req_list = QListWidget()
    try:
        self.req_list.setWordWrap(False)
    except Exception:
        pass

    self.req_list.itemSelectionChanged.connect(self._on_select_row)
    bottom_v.addWidget(self.req_list, 1)
    self.splitter.addWidget(bottom_wrap)
    self.splitter.setSizes([self.height() * 3 // 4, self.height() // 4])

    root.addWidget(self.splitter, 1)

    try:
        QShortcut(QKeySequence.MoveToPreviousChar, self, activated=self.go_prev_page)  
        QShortcut(QKeySequence.MoveToNextChar, self, activated=self.go_next_page)      
        QShortcut(QKeySequence.MoveToPreviousPage, self, activated=self.go_prev_page)
        QShortcut(QKeySequence.MoveToNextPage, self, activated=self.go_next_page)    
    except Exception:
        pass

    def _on_filter_changed(self, idx: int):
        self.filter_mode = ["all", "found", "missing", "manual"][idx]
        self.selected_row_id = None
        self._refresh_sidebar()
        self.render_page()

    def _row_status_icon(self, status: str) -> str:
        return {"found": "✅", "missing": "❌", "manual": "🛠️"}.get(status, "•")

    def _flatten_exact_wording(self, text: str) -> str:
        text = (text or "").replace("\r", "")
        parts = [p.strip() for p in text.split("\n") if p.strip()]
        return " | ".join(parts) if parts else "-"

    def _fmt_badge(self, r: Dict) -> str:
        keys = ("format", "Format", "style", "Style", "Bold", "Underline")
        found = []
        for k in keys:
            v = r.get(k)
            if v:
                vs = str(v).strip()
                if vs and vs not in found:
                    found.append(vs)
        return ("  |  Format: " + ", ".join(found)) if found else ""

    def _refresh_sidebar(self):
        self.req_list.clear()
        total = len(self.rows)
        frows = self._filtered_rows()
        for idx, r in enumerate(frows, start=1):
            icon = self._row_status_icon(r.get("status", ""))
            pages_spec = r.get("pages_spec") or "-"
            exact_disp = self._flatten_exact_wording(r.get("symbol", "-"))
            text = f"{idx}. {icon} {r.get('requirement','-')}  |  Exact wording: {exact_disp}  |  Pages: {pages_spec}{self._fmt_badge(r)}"
            item = QListWidgetItem(text)
            item.setData(Qt.UserRole, r.get("id"))
            self.req_list.addItem(item)
        self.info_label.setText(f"Requirements: {total}  |  Showing: {len(frows)}")

    def _filtered_rows(self) -> List[Dict]:
        if self.filter_mode == "all":
            return self.rows
        return [r for r in self.rows if r.get("status") == self.filter_mode]

    def _on_select_row(self):
        items = self.req_list.selectedItems()
        self.selected_row_id = items[0].data(Qt.UserRole) if items else None
        self.render_page()

    # -------------------------- Terms & Cache --------------------------
    def _active_rows_for_page(self, page_no: int) -> List[Dict]:
        candidates = self._filtered_rows()
        out = []
        for r in candidates:
            pset = parse_pages_spec(r.get("pages_spec", ""), self.page_count)
            if (pset is None) or (page_no in pset):
                out.append(r)
        return out

    def _active_terms_for_page(self, page_no: int) -> List[str]:
        rows = self._active_rows_for_page(page_no)
        if self.selected_row_id is not None:
            rows = [r for r in rows if r.get("id") == self.selected_row_id]
        terms: List[str] = []
        for r in rows:
            terms.extend(build_terms_from_symbol(r.get("symbol", "")))
        seen = set(); out=[]
        for t in terms:
            if t not in seen:
                seen.add(t); out.append(t)
        return out[:200]

    def _get_cache(self, page_no: int) -> Dict[str, List[fitz.Rect]]:
        if page_no in self._cache:
            v = self._cache.pop(page_no)
            self._cache[page_no] = v
            return v
        if len(self._cache) >= self._lru_capacity:
            self._cache.popitem(last=False)
        self._cache[page_no] = {}
        return self._cache[page_no]

    def _search_term_on_page(self, page: fitz.Page, term: str) -> List[fitz.Rect]:
        try:
            return page.search_for(term) or []
        except Exception:
            return []

    def _hits_for_page(self, page_no: int, terms: List[str]) -> List[fitz.Rect]:
        cache = self._get_cache(page_no)
        page = self.doc.load_page(page_no)
        page_rect = page.rect
        artwork_rect = shrink_rect(
            page_rect, ARTWORK_MARGIN_L, ARTWORK_MARGIN_R, ARTWORK_MARGIN_T, ARTWORK_MARGIN_B
        )

        out: List[fitz.Rect] = []
        for t in terms:
            if t not in cache:
                cache[t] = self._search_term_on_page(page, t)
            for r in cache[t]:
                if rect_center_inside(r, artwork_rect):
                    out.append(r)

        if not out:
            return out
        
        cell = 8.0
        buckets = {}
        uniq = []
        for r in out:
            key = (int(r.x0 // cell), int(r.y0 // cell), int(r.x1 // cell), int(r.y1 // cell))
            if key not in buckets:
                buckets[key] = True
                uniq.append(r)
        return uniq

    # -------------------------- Nav & Zoom --------------------------
    def _update_nav_state(self):
        self.page_label.setText(f"Page {self.current_page + 1} / {self.page_count}")
        self.prev_page_btn.setEnabled(self.current_page > 0)
        self.next_page_btn.setEnabled(self.current_page < self.page_count - 1)

    def go_prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.render_page()
            QTimer.singleShot(0, self._prefetch_neighbors)

    def go_next_page(self):
        if self.current_page < self.page_count - 1:
            self.current_page += 1
            self.render_page()
            QTimer.singleShot(0, self._prefetch_neighbors)

    def set_zoom(self, z: float):
        self.zoom = max(0.2, min(6.0, float(z)))
        self.render_page()

    def fit_width(self):
        try:
            page = self.doc.load_page(self.current_page)
            pm = page.get_pixmap(matrix=fitz.Matrix(1, 1), alpha=False)
        except Exception:
            return
        if pm.width <= 0:
            return
        viewport_w = max(50, self.view.viewport().width() - 16)
        self.zoom = max(0.2, min(6.0, viewport_w / pm.width))

    def _prefetch_neighbors(self):
        for delta in (-1, 1):
            p = self.current_page + delta
            if 0 <= p < self.page_count:
                terms = self._active_terms_for_page(p)
                if terms:
                    self._hits_for_page(p, terms)

    # -------------------------- Render --------------------------
    def render_page(self):
        if self._render_busy:
            self._render_again = True
            return
        self._render_busy = True
        try:
            self.scene.clear()
            self._update_nav_state()
            self._syncing_pagebar = True
            try:
                self.page_bar.setRange(1, max(1, self.page_count))
                self.page_bar.setValue(self.current_page + 1)
            finally:
                self._syncing_pagebar = False

            page = self.doc.load_page(self.current_page)
            matrix = fitz.Matrix(self.zoom, self.zoom)
            pm = page.get_pixmap(matrix=matrix, alpha=False)
            pixmap = QPixmap()
            pixmap.loadFromData(pm.tobytes("png"))
            pix_item = self.scene.addPixmap(pixmap)
            self.scene.setSceneRect(QRectF(pixmap.rect()))


            terms = self._active_terms_for_page(self.current_page)
            rects = self._hits_for_page(self.current_page, terms) if terms else []
            if rects:
                pen = QPen(Qt.NoPen)
                brush = QBrush(QColor(255, 235, 59, HIGHLIGHT_ALPHA)) 
                for r in rects:
                    rr = QRectF(r.x0 * self.zoom, r.y0 * self.zoom,
                                (r.x1 - r.x0) * self.zoom, (r.y1 - r.y0) * self.zoom)
                    self.scene.addRect(rr, pen, brush)

            self.view.centerOn(pix_item)
        finally:
            self._render_busy = False

        if self._render_again:
            self._render_again = False
            QTimer.singleShot(0, self.render_page)

    def _draw_artwork_box(self, page: fitz.Page, pm: fitz.Pixmap):
        page_rect = page.rect
        box = shrink_rect(page_rect, ARTWORK_MARGIN_L, ARTWORK_MARGIN_R, ARTWORK_MARGIN_T, ARTWORK_MARGIN_B)
        pen_box = QPen(QColor(66, 133, 244, 140), 2)  # ฟ้าโปร่ง
        rr = QRectF(box.x0 * self.zoom, box.y0 * self.zoom,
                    (box.x1 - box.x0) * self.zoom, (box.y1 - box.y0) * self.zoom)
        self.scene.addRect(rr, pen_box, QtGui.QBrush(Qt.NoBrush))

    def closeEvent(self, e: QtGui.QCloseEvent):
        try:
            self.doc.close()
        except Exception:
            pass
        super().closeEvent(e)

# ------------------------------ Top-level Window ------------------------------
class PdfPreviewWindow(QMainWindow):
    def __init__(self, pdf_path: str, rows: list, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Preview PDF")
        self.setAttribute(Qt.WA_DeleteOnClose, True)
        self.resize(1200, 820)
        self.viewer = PDFViewer(pdf_path=pdf_path, rows=rows, parent=None)
        self.setCentralWidget(self.viewer)

        QShortcut(QKeySequence.ZoomIn,  self, activated=lambda: self.viewer.set_zoom(self.viewer.zoom * 1.1))
        QShortcut(QKeySequence.ZoomOut, self, activated=lambda: self.viewer.set_zoom(self.viewer.zoom * 0.9))
        QtCore.QTimer.singleShot(0, self.viewer.fit_width)
