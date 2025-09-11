import os, logging
import re 
import pandas as pd
import unicodedata as _ud
import html as _html
import html
from PyQt5.QtWidgets import QWidget, QLabel, QHBoxLayout, QVBoxLayout, QGridLayout, QHeaderView, QTableWidgetSelectionRange
from PyQt5.QtCore import Qt, QUrl
from PyQt5 import QtWidgets, QtGui, QtCore
from ui.pdf_viewer import PdfPreviewWindow
from checklist_loader import load_checklist, start_check, extract_part_code_from_pdf
from pdf_reader import extract_text_by_page
from checker import check_term_in_page
from result_exporter import export_result_to_excel
from PyQt5.QtGui import QColor, QIcon, QPixmap, QDesktopServices
from pdf_reader import extract_product_info_by_page
from collections import defaultdict



APP_ICON_PATH = os.path.join("assets", "app", "dso_icon.ico")

# OCR Language sets
LANG_SUPERSET = "eng+spa+fra+por+ita+deu+nld+swe+fin+dan+nor+pol+ces+slk+hun+rus+ell+tur+ara+jpn+chi_sim+tha"

PART_OCR_MAP_FAST = {
    "4LB":  "eng+spa+fra+por",
    "DOM":      "eng",
    "UU1":      "eng",
    "UU1_DOM":  "eng",
    "21A": "eng+spa+fra+por+tha",
    "19A": "eng+spa+fra+por+tha",
    "19L": "eng+spa+fra+por+tha",
    "2LB": "eng+fra",
    "2XV": "chi_sim+eng",
    "DC1": "eng+fra+deu+ita+nld+spa+por",
}

# สเปกภาษาของแต่ละ Part 
PART_OCR_MAP_FULL = {
    "4LB":  "eng+spa+fra+por",
    "DOM":  "eng",
    "UU1":  "eng",
    "UU1_DOM": "eng",
    "21A": "eng+fra+deu+ita+nld+spa+por+swe+fin+dan+nor+rus+pol+ces+slk+hun",
    "19A": "eng+fra+deu+ita+nld+spa+por+swe+fin+dan+nor+pol+ces+slk+hun+rus+ell+tur+ara",
    "19L": "eng+fra+deu+ita+nld+spa+por+swe+fin+dan+nor+pol+ces+slk+hun+rus+ell+tur+ara",
    "2LB": "eng+fra",
    "2XV": "chi_sim+eng",
    "DC1": "eng+fra+deu+ita+nld+spa+por+pol+ces+hun+jpn",
}

# ==== UI-only: hide entire SPW/SPG group if the whole group is "Not Found" (robust by Requirement) ====
def _hide_empty_sp_group_ui(df):
    """
    ซ่อนทั้ง 'กลุ่ม' ของ International warning statement:
      - กลุ่ม SPW หรือ SPG จะถูก 'ซ่อนทั้งกลุ่ม' ถ้าในกลุ่มนั้นไม่มีแถวใด Found เลย
    เงื่อนไขการจับกลุ่ม:
      - อ่านจากคอลัมน์ Requirement เป็นหลัก (เพราะไฟล์ของคุณระบุ ': SPW' / ': SPG' ที่นี่)
      - ทนต่อช่องว่างซ้ำ, โคลอนหลายแบบ (:/：), และขีด ( - / – / — )
    เกณฑ์ Found:
      - ถ้า 'Found' เริ่มด้วย '✅' = found
      - ถ้าไม่มี 'Found' ให้ fallback ใช้ Pages (ไม่ใช่ '-', '—', '', 'none', '0')
    """
    try:
        import pandas as pd

        if df is None or df.empty:
            return df

        cols = set(df.columns)
        if "Requirement" not in cols:
            return df

        REQ_COL   = "Requirement"
        PAGES_COL = "Pages"
        FOUND_COL = "Found" if "Found" in cols else None

        def _norm(s: str) -> str:
            # normalize: NFKC, lower, collapse spaces, unify dashes
            s = _ud.normalize("NFKC", str(s or ""))
            s = s.replace("\u2013", "-").replace("\u2014", "-").replace("\u2212", "-")
            s = s.lower()
            s = re.sub(r"\s+", " ", s).strip()
            return s

        # regex: international\s+warning\s+statement\s*[:：\-]?\s*(spw|spg)\b
        PAT = re.compile(r"international\s+warning\s+statement\s*[:：\-]?\s*(spw|spg)\b", re.I)

        def _row_tag(row) -> str:
            req = _norm(row.get(REQ_COL, ""))
            m = PAT.search(req)
            if not m:
                return ""   # ไม่ใช่แถวของ SPW/SPG
            return m.group(1).lower()  # 'spw' หรือ 'spg'

        def _pages_not_empty(v) -> bool:
            if v is None: return False
            if isinstance(v, (set, list, tuple)): return len(v) > 0
            s = str(v).strip().lower()
            return s not in ("", "-", "—", "none", "0")

        def _row_found(row) -> bool:
            if FOUND_COL:
                f = str(row.get(FOUND_COL, "")).strip()
                if f.startswith("✅"):
                    return True
                if f.startswith("❌"):
                    return False
            return _pages_not_empty(row.get(PAGES_COL, None))

        # ทำ tagging ทีละแถวจาก Requirement
        tags = [ _row_tag(df.iloc[i]) for i in range(len(df)) ]
        if not any(t in ("spw","spg") for t in tags):
            return df  # ไม่มีแถว SP เลย

        is_spw = [t == "spw" for t in tags]
        is_spg = [t == "spg" for t in tags]

        spw_found_any = any(_row_found(df.iloc[i]) for i, m in enumerate(is_spw) if m)
        spg_found_any = any(_row_found(df.iloc[i]) for i, m in enumerate(is_spg) if m)

        keep_idx = []
        for i in range(len(df)):
            if is_spw[i] and not spw_found_any:
                continue  # ซ่อนทั้งกลุ่ม SPW
            if is_spg[i] and not spg_found_any:
                continue  # ซ่อนทั้งกลุ่ม SPG
            keep_idx.append(i)

        return df.iloc[keep_idx].copy() if keep_idx else df.iloc[0:0].copy()

    except Exception as e:
        try:
            print("[UI-HIDE-SP] error:", e)
        except:
            pass
        return df

# ==== Preview helpers for SP rules ====
def _prune_spw_prefix_terms_if_spg_present(terms, df_result):
    try:
        if df_result is None:
            return terms

        REQ_COL   = "Requirement"
        PAGES_COL = "Pages"
        if not (isinstance(df_result, type(None)) or df_result.empty):
            def _pages_not_empty(v):
                if v is None: return False
                if isinstance(v, (set, list, tuple)): return len(v) > 0
                s = str(v).strip()
                return not (s == "" or s == "-" or s == "—" or s.lower() == "none" or s == "0")

            df_found = df_result[df_result[PAGES_COL].map(_pages_not_empty)] if (PAGES_COL in df_result.columns) else df_result
            if REQ_COL in df_found.columns:
                req_norm = df_found[REQ_COL].fillna("").str.lower()
                have_spg_found = bool((req_norm.str.contains("international warning statement", regex=False)
                                       & req_norm.str.contains(r"\bspg\b", regex=True)).any())
                if have_spg_found:
                    def _is_short_spw(s: str) -> bool:
                        s2 = (s or "").lower()
                        return ("warning" in s2 and "small parts" in s2 and "may be generat" not in s2)
                    terms = [t for t in terms if not _is_short_spw(t)]
        return terms
    except Exception:
        return terms

def _get_ocr_langs_for_part(part_code: str):
    code = (part_code or "").strip().upper()
    fast = PART_OCR_MAP_FAST.get(code, "eng")
    full = PART_OCR_MAP_FULL.get(code, fast)
    return fast, full

class _PdfWorker(QtCore.QThread):
    finished = QtCore.pyqtSignal(list, list) 
    error = QtCore.pyqtSignal(str)

    def __init__(self, path):
        super().__init__()
        self.path = path

    def run(self):
        try:
            try:
                codes = extract_part_code_from_pdf(self.path) or []
            except Exception:
                codes = []
            part_code = (codes[0] if codes else getattr(self, "part_code", "")) or ""

            fast_lang, full_lang = _get_ocr_langs_for_part(part_code)

            pages = extract_text_by_page(
                self.path,
                enable_ocr=True,
                ocr_only_suspect_pages=True,   
                ocr_lang_fast=fast_lang,        
                ocr_lang_full=full_lang         
            )
            infos = extract_product_info_by_page(pages)
            self.finished.emit(pages, infos)

        except Exception as e:
            self.finished.emit([], {"error": str(e)})

class _ExcelWorker(QtCore.QThread):
    finished = QtCore.pyqtSignal(object)     
    error = QtCore.pyqtSignal(str)
    def __init__(self, path: str, pdf_basename: str):
        super().__init__()
        self.path = path
        self.pdf_basename = pdf_basename
    def run(self):
        try:
            df = load_checklist(self.path, self.pdf_basename)
            self.finished.emit(df)
        except Exception as e:
            self.error.emit(str(e))

class _CheckWorker(QtCore.QThread):
    finished = QtCore.pyqtSignal(object)
    error = QtCore.pyqtSignal(str)
    def __init__(self, df_checklist, pages):
        super().__init__()
        self.df_checklist = df_checklist
        self.pages = pages
    def run(self):
        try:
            res = start_check(self.df_checklist, self.pages)
            self.finished.emit(res)
        except Exception as e:
            self.error.emit(str(e))

RED_HEX = "#ff1313"
Y_ROW    = "#FFFACD"  
Y_HOVER  = "#FCF4AF"  
Y_SEL    = "#fff1b0"  
G_HOVER  = "#F9F9F9"
G_SEL    = "#dddddd"
SEL_ROW = "#E6F2FF" 


# Image sizing rules
LOGO_MAX_WIDTH_PX = 170       
IMG_SIDE_PADDING  = 8        
LOGO_KEYS = (" logo", " mark", " lion", " ce ", " ukca", " mc ", "cib")
FORCE_FULL_KEYS = ("warning", "statement", "spw", "international",
                    "upc", "list of content", "address", "instruction")

def _is_logo_name(path: str, req_text: str) -> bool:
    s = (os.path.basename(path or "") + " " + (req_text or "")).lower()
    s = " " + s.replace("_", " ") + " "
    return any(k in s for k in LOGO_KEYS)

def _must_fill_width(req_text: str) -> bool:
    s = " " + (req_text or "").lower() + " "
    return any(k in s for k in FORCE_FULL_KEYS)

class DSOApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        try:
            self.setWindowIcon(QIcon(APP_ICON_PATH))
        except Exception:
            pass

        self.setWindowTitle("DSO - Digital Sign Off")
        self.setGeometry(100, 100, 1200, 800)

        self.excel_path = ""
        self.pdf_path = ""
        self.checklist_df = None
        self.pages = None
        self.result_df = None
        self.product_infos = []
        self._image_cache = {}

        self.init_ui()

    def init_ui(self):
        shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+F"), self)
        shortcut.activated.connect(self.search_text)
        layout = QtWidgets.QVBoxLayout()

        # Upload buttons
        file_layout = QtWidgets.QHBoxLayout()
        self.pdf_btn = QtWidgets.QPushButton("📄 Upload PDF (Artwork)")
        self.excel_btn = QtWidgets.QPushButton("📋 Upload Excel (Checklist)")
        self.pdf_label = QtWidgets.QLabel("PDF: Not selected")
        self.excel_label = QtWidgets.QLabel("Excel: Not selected")
        self.pdf_btn.clicked.connect(self.load_pdf)
        self.excel_btn.clicked.connect(self.load_excel)
        file_layout.addWidget(self.pdf_btn)
        file_layout.addWidget(self.excel_btn)
        file_layout.addStretch()

        # Search bar
        search_layout = QtWidgets.QHBoxLayout()
        self.search_input = QtWidgets.QLineEdit()
        self.search_input.returnPressed.connect(self.search_text)
        self.search_input.textChanged.connect(self._on_search_text_changed)
        self.search_input.setPlaceholderText("Search term ...")
        self.search_btn = QtWidgets.QPushButton("Search")
        self.search_btn.clicked.connect(self.search_text)
        search_layout.addWidget(self.search_input)
        search_layout.addWidget(self.search_btn)

        # Result Table
        self.result_table = QtWidgets.QTableWidget()
        self.result_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.result_table.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        self.result_table.setAlternatingRowColors(True)
        self.result_table.setColumnCount(0)
        self.result_table.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.result_table.setWordWrap(True)
        self.result_table.verticalHeader().setSectionResizeMode(QtWidgets.QHeaderView.ResizeToContents)
        self.result_table.horizontalHeader().setStretchLastSection(False)
        self.result_table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Interactive)
        self.result_table.setShowGrid(True)
        self.result_table.setProperty("searchMode", False)
        self.result_table.setStyleSheet(
            (self.result_table.styleSheet() or "") + """
        QTableView { gridline-color: #BFBFBF; }

        QTableWidget[searchMode="true"]::item:selected,
        QTableView[searchMode="true"]::item:selected {
            background: transparent;
            color: inherit;
        }

        QTableWidget::item:focus, QTableView::item:focus { outline: none; }
        """
        )
        pal = self.result_table.palette()
        pal.setColor(QtGui.QPalette.Highlight, QtGui.QColor(SEL_ROW))        
        pal.setColor(QtGui.QPalette.HighlightedText, QtGui.QColor(Qt.black)) 
        self.result_table.setPalette(pal)
        self.hovered_row = -1
        self.result_table.setMouseTracking(True)
        self.result_table.cellClicked.connect(self._on_table_cell_clicked)
        self.result_table.cellDoubleClicked.connect(self._on_table_cell_double_clicked)
        self.result_table.viewport().installEventFilter(self)
        self.result_table.itemSelectionChanged.connect(self.update_row_highlight)
        self.result_table.horizontalHeader().sectionResized.connect(self._on_column_resized)

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

    def eventFilter(self, source, event):
        if source == self.result_table.viewport() and event.type() == QtCore.QEvent.MouseMove:
            index = self.result_table.indexAt(event.pos())
            row = index.row()
            if row != self.hovered_row:
                self.hovered_row = row
                self.update_row_highlight()
        return super().eventFilter(source, event)
    
    def _set_hide_selection_overlay_during_search(self, hide: bool):
        self.result_table.setProperty("searchMode", bool(hide))
        self.result_table.style().unpolish(self.result_table)
        self.result_table.style().polish(self.result_table)
        self.result_table.viewport().update()

    def _select_entire_row(self, row: int):
        if row < 0 or row >= self.result_table.rowCount():
            return
        sm = self.result_table.selectionModel()
        if sm:
            sm.clearSelection()
            left  = self.result_table.model().index(row, 0)
            right = self.result_table.model().index(row, self.result_table.columnCount()-1)
            sel = QtCore.QItemSelection(left, right)
            sm.select(sel, QtCore.QItemSelectionModel.ClearAndSelect | QtCore.QItemSelectionModel.Rows)
            self.result_table.setCurrentIndex(left)
    
    class _RowClickFilter(QtCore.QObject):
        def __init__(self, table, row, parent=None):
            super().__init__(parent)
            self._table = table
            self._row = row
        def eventFilter(self, obj, event):
            if event.type() == QtCore.QEvent.MouseButtonPress:
                if isinstance(obj, QtWidgets.QLabel):
                    try:
                        t = (obj.text() or "").lower()
                    except Exception:
                        t = ""
                    if "<a " in t or 'href="' in t:
                        return False
                self._table.setCurrentCell(self._row, 0)
                self._table.selectRow(self._row)
                return False
            return False

    def _attach_row_select(self, widget: QtWidgets.QWidget, row_idx: int):
        f = DSOApp._RowClickFilter(self.result_table, row_idx, self)
        widget.installEventFilter(f)
        for ch in widget.findChildren(QtWidgets.QWidget):
            if isinstance(ch, QtWidgets.QLabel):
                continue
            ch.installEventFilter(f)
        if not hasattr(self, "_row_filters"):
            self._row_filters = []
        self._row_filters.append(f)

    def _autosize_column_to_contents(self, column_name: str, min_w: int = 300, max_w: int = 700):
        col_idx = self.get_column_index(column_name)
        if col_idx < 0:
            return
        header = self.result_table.horizontalHeader()
        need = min_w
        for row in range(self.result_table.rowCount()):
            w = self.result_table.cellWidget(row, col_idx)
            if w:
                for lbl in w.findChildren(QtWidgets.QLabel):
                    txt = re.sub(r"<[^>]+>", "", lbl.text() or "")
                    fm = lbl.fontMetrics()
                    width = fm.boundingRect(txt).width() + 24
                    need = max(need, width)

            it = self.result_table.item(row, col_idx)
            if it:
                fm = it.fontMetrics()
                width = fm.boundingRect(it.text() or "").width() + 24
                need = max(need, width)
        need = max(min_w, min(int(need), max_w))
        header.setSectionResizeMode(col_idx, QtWidgets.QHeaderView.Fixed)
        self.result_table.setColumnWidth(col_idx, need)

    def update_row_highlight(self):
        verif_col = self.get_column_index("Verification")

        sel_hex = SEL_ROW

        def paint_cell(row, col, bg_hex, is_selected):
            it = self.result_table.item(row, col)
            if it:
                if is_selected:
                    it.setBackground(QtGui.QBrush())
                else:
                    it.setBackground(QColor(bg_hex))

            w = self.result_table.cellWidget(row, col)
            if w:
                w.setStyleSheet(
                    f"QWidget{{background-color:{bg_hex};}}"
                    f" QLabel{{background-color:{bg_hex};}}"
                )
                for ch in w.findChildren(QtWidgets.QLabel):
                    prev = ch.styleSheet() or ""
                    if "background-color" in prev:
                        ch.setStyleSheet(re.sub(r"background-color:[^;]+;", f"background-color:{bg_hex};", prev))
                    else:
                        ch.setStyleSheet(prev + f"background-color:{bg_hex};")

        for row in range(self.result_table.rowCount()):
            is_selected = self.result_table.selectionModel().isRowSelected(row, QtCore.QModelIndex())
            is_hovered  = (row == self.hovered_row)

            is_manual = False
            if verif_col != -1:
                verif_item = self.result_table.item(row, verif_col)
                if verif_item:
                    tag = verif_item.data(QtCore.Qt.UserRole)
                    if isinstance(tag, str) and tag.strip().lower() == "manual":
                        is_manual = True
                    elif isinstance(verif_item.text(), str) and verif_item.text().strip().lower() == "manual":
                        is_manual = True

            # กำหนดพื้นหลังฐานของทั้งแถว
            if is_selected:
                base_bg = sel_hex
            else:
                if is_manual:
                    base_bg = Y_HOVER if is_hovered else Y_ROW 
                else:
                    base_bg = G_HOVER if is_hovered else "white"

            for col in range(self.result_table.columnCount()):
                cell_bg = base_bg
                paint_cell(row, col, cell_bg, is_selected)

    def get_column_index(self, column_name):
        for col in range(self.result_table.columnCount()):
            header_item = self.result_table.horizontalHeaderItem(col)
            if header_item and header_item.text().strip().lower() == column_name.strip().lower():
                return col
        return -1
    
    def _collect_symbol_labels(self):
        sym_idx = self.get_column_index("Symbol/ Exact wording")
        if sym_idx < 0:
            return []
        out = []
        for r in range(self.result_table.rowCount()):
            cell = self.result_table.cellWidget(r, sym_idx)
            if not cell:
                continue
            for lbl in cell.findChildren(QtWidgets.QLabel):
                base = lbl.property("base_inner_html")
                if isinstance(base, str):
                    out.append((r, sym_idx, lbl))
        return out

    def _set_label_inner_html(self, lbl: QtWidgets.QLabel, inner_html: str):
        lbl.setTextFormat(QtCore.Qt.RichText)
        lbl.setText(self._wrap_html_with_table_font(inner_html or ""))

    def _clear_symbol_highlight(self):
        for _, _, lbl in self._collect_symbol_labels():
            base = lbl.property("base_inner_html")
            if isinstance(base, str):
                self._set_label_inner_html(lbl, base)

    def _apply_symbol_highlight(self, query: str) -> int:
        if not query:
            self._clear_symbol_highlight()
            return -1

        pat = re.compile(re.escape(query), flags=re.IGNORECASE)
        first_row = -1

        for r, c, lbl in self._collect_symbol_labels():
            base = lbl.property("base_inner_html") or ""
            def repl(m):
                return f'<span style="background:#ffeb3b">{m.group(0)}</span>'
            highlighted = pat.sub(repl, base)

            if highlighted != base and first_row == -1:
                first_row = r
            self._set_label_inner_html(lbl, highlighted)

        return first_row

    def _on_search_text_changed(self, s: str):
        text = (s or "").strip()
        self._set_hide_selection_overlay_during_search(True)
        self._apply_symbol_highlight(text)
        if self.result_table.selectionModel():
            self.result_table.selectionModel().clearSelection()
    
    def _table_font_pt(self) -> int:
        pt = self.result_table.font().pointSize()
        return max(8, pt if pt > 0 else 10)

    def _wrap_html_with_table_font(self, body_html: str) -> str:
        f = self.result_table.font()
        fam  = f.family()
        size = f.pointSizeF() if f.pointSizeF() > 0 else float(f.pointSize() or 10)

        css = f"""
        <style>
        html, body {{
            margin:0; padding:0;
            font-family: "{fam}";
            font-size: {size:.2f}pt;
            color: #000;
        }}
        p, div, span {{
            font-family: "{fam}";
            font-size: {size:.2f}pt;
            color: #000;
        }}
        b, strong {{ font-weight: bold !important; }}
        u {{ text-decoration: underline !important; }}
        a {{
            color: #1a73e8;
            text-decoration: underline;
        }}
        </style>
        """
        return f"<html><head>{css}</head><body>{body_html or ''}</body></html>"
    
    def _wrap_all_as_link(self, inner_html: str, url: str) -> str:
        if not url:
            return inner_html
        u = url.strip()
        if u.lower().startswith("www."):
            u = "http://" + u
        return (
            f'<a href="{_html.escape(u)}" '
            f'style="color:#1a73e8; text-decoration: underline;">{inner_html}</a>'
        )
    
    def _linkify_plain_urls(self, html_text: str) -> str:
        def anchor(u: str, label: str = None) -> str:
            if not label:
                label = u
            return (
                f'<a href="{u}" '
                f'style="color:#1a73e8; text-decoration: underline;">{label}</a>'
            )

        html_text = re.sub(
            r'(?<!href=")(?P<url>https?://[^\s<>"\')]+)',
            lambda m: anchor(m.group("url")),
            html_text,
        )

        html_text = re.sub(
            r'(?<!href=")(?P<url>www\.[^\s<>"\')]+)',
            lambda m: anchor("http://" + m.group("url"), m.group("url")),
            html_text,
        )
        return html_text
    
    def _remark_pairs_to_html(self, text: str, inner_w: int, fm: QtGui.QFontMetrics) -> str:
        esc = _html.escape
        t = str(text or "").replace("\r", "\n").strip()
        t_no_ws = re.sub(r"[ \t\f\v\u00A0]", "", t)
        t_no_ws = t_no_ws.replace("–", "-").replace("—", "-").replace("=", "")
        if t_no_ws == "" or set(t_no_ws) <= {"-"}:            
            return "-"

        # จับคู่ทั่วไป "LEFT = RIGHT" (ยืดหยุ่น ครอบคลุมอักขระยุโรป)
        PAIR_FULL = re.compile(r"\s*([^=]+?)\s*=\s*([^\n=]+)\s*$")

        def fmt_pair(left: str, right: str) -> str:
            left, right = left.strip(), right.strip()
            one_line = f"{left} = {right}"
            if fm.horizontalAdvance(one_line) >= int(inner_w * 0.92):
                return f"{esc(left)} =<br/>{esc(right)}"
            return esc(one_line)

        # กรณีมี \n อยู่แล้ว → ประมวลผลทีละบรรทัด
        if "\n" in t:
            lines = []
            for line in t.split("\n"):
                line = line.strip()
                if not line:
                    continue
                m = PAIR_FULL.fullmatch(line)
                lines.append(fmt_pair(*m.groups()) if m else esc(line))
            return "<br/>".join(lines) if lines else "-"

        # สตริงเดียว: อาจมีหลายคู่ในบรรทัดเดียว
        pairs = [(m.group(1), m.group(2)) for m in re.finditer(r"([^=]+?)\s*=\s*([^\n=]+)", t)]
        if len(pairs) >= 2:
            return "<br/>".join(fmt_pair(l, r) for (l, r) in pairs)

        m = PAIR_FULL.fullmatch(t)
        if m:
            return fmt_pair(m.group(1), m.group(2))
        return esc(t)
    
    def _pairs_to_multiline_html(self, s: str) -> str:

        t = str(s or "").replace("\r", "\n")
        if "\n" in t:
            return _html.escape(t).replace("\n", "<br/>")

        # จับคู่ข้อความตัวพิมพ์ใหญ่ (รวมอักขระยุโรป) ตามด้วย " = " และชื่อภาษา
        PAIR_RE = re.compile(
            r"([A-Z0-9\u00C0-\u017F][A-Z0-9\s'’\-\u00C0-\u017F]+?\s*=\s*[A-Za-z\u00C0-\u017F]+)"
        )
        pairs = PAIR_RE.findall(t)
        if len(pairs) >= 2:
            return "<br/>".join(_html.escape(p.strip()) for p in pairs)

        return _html.escape(t)

    def load_pdf(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Select PDF Artwork", "", "PDF Files (*.pdf)")
        if not path:
            return

        self.pdf_path = path
        self.pdf_label.setText(f"PDF: {os.path.basename(path)}")

        # กันกดซ้ำระหว่างโหลด
        self.pdf_btn.setEnabled(False)
        self.excel_btn.setEnabled(False)
        self.check_btn.setEnabled(False)

        self._pdf_worker = _PdfWorker(self.pdf_path)

        def _ok(pages, infos):
            self.pages = pages
            self.product_infos = infos or []
            self.pdf_btn.setEnabled(True)
            self.excel_btn.setEnabled(True)
            self.check_btn.setEnabled(bool(self.checklist_df))

        def _err(msg):
            QtWidgets.QMessageBox.critical(self, "PDF Error", msg)
            self.pdf_btn.setEnabled(True)
            self.excel_btn.setEnabled(True)
            self.check_btn.setEnabled(bool(self.checklist_df))

        self._pdf_worker.finished.connect(_ok)
        self._pdf_worker.error.connect(_err)
        self._pdf_worker.start()

    def load_excel(self):
        if not getattr(self, "pdf_path", None):
            QtWidgets.QMessageBox.warning(self, "PDF Required", "Please upload a PDF file first.")
            return

        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Select Checklist Excel", "", "Excel Files (*.xlsx *.xls)")
        if not path:
            return

        self.excel_path = path
        self.excel_label.setText(f"Checklist: {os.path.basename(path)}")

        self.pdf_btn.setEnabled(False)
        self.excel_btn.setEnabled(False)
        self.check_btn.setEnabled(False)

        self._excel_worker = _ExcelWorker(self.excel_path, os.path.basename(self.pdf_path))

        def _ok(df):
            self.checklist_df = df
            self.pdf_btn.setEnabled(True)
            self.excel_btn.setEnabled(True)
            self.check_btn.setEnabled(bool(self.pages))

        def _err(msg):
            QtWidgets.QMessageBox.critical(self, "Checklist Load Error", msg)
            self.pdf_btn.setEnabled(True)
            self.excel_btn.setEnabled(True)
            self.check_btn.setEnabled(bool(self.pages))

        self._excel_worker.finished.connect(_ok)
        self._excel_worker.error.connect(_err)
        self._excel_worker.start()

    def start_checking(self):
        if getattr(self, "checklist_df", None) is None or not getattr(self, "pages", None):
            QtWidgets.QMessageBox.warning(self, "Missing File", "Please upload both Checklist and PDF before checking.")
            return

        self.check_btn.setEnabled(False)
        self.export_btn.setEnabled(False)

        self._check_worker = _CheckWorker(self.checklist_df, self.pages)

        def _ok(df):
            self.result_df = df
            if not isinstance(self.result_df, pd.DataFrame) or self.result_df.empty:
                QtWidgets.QMessageBox.information(self, "No Result", "No matching terms found.")
            else:
                self.display_results(self.result_df)
            self.check_btn.setEnabled(True)
            self.export_btn.setEnabled(True)

        def _err(msg):
            QtWidgets.QMessageBox.critical(self, "Check Error", msg)
            self.check_btn.setEnabled(True)
            self.export_btn.setEnabled(True)

        self._check_worker.finished.connect(_ok)
        self._check_worker.error.connect(_err)
        self._check_worker.start()

    def display_results(self, df: pd.DataFrame):
        df_src = df.copy()

        symbol_cols_protect = {"Symbol/ Exact wording", "Symbol/Exact wording", "__Term_HTML__"}
        cols_to_fill = [c for c in df.columns if c not in symbol_cols_protect]
        df[cols_to_fill] = df[cols_to_fill].fillna("-")

        def _dashify(x):
            s = str(x).strip()
            if s == "":
                return "-"
            if s.lower() in {"none", "nan", "null"}:
                return "-"
            if re.fullmatch(r"[-–—=\s]+", s):
                return "-"
            return s

        for col in ["Found", "Match", "Font Size", "Pages", "Note", "Remark", "Package Panel", "Procedure"]:
            if col in df.columns:
                df[col] = df[col].apply(_dashify)

        preferred = ["Verification", "Requirement", "Symbol/ Exact wording", "Specification", 
                    "Found", "Match", "Font Size", "Pages", "Note", 
                    "Package Panel", "Procedure", "Remark" ]
        
        # helper column ที่ไม่แสดงใน UI แต่ยังแสดงใน df_scr
        helper_names = {"remark url", "remark link"}
        helper_cols = [ c for c in df.columns if c.strip().lower() in helper_names]

        # คอลัมน์ภายในที่ต้องซ่อนจาก UI
        internal_hide = {
            "__Term_HTML__", "Image_Groups_Resolved", "Image_Groups",
            "Image Path Resolved", "Image Path", "_HasImage", "Language List"
        }

        # จัดลำดับเฉพาะคอลัมน์ที่จะแสดง (ไม่รวม helper)
        ordered = [ c for c in preferred if c in df.columns]
        tail = [c for c in df.columns if c not in ordered + helper_cols]

        # ตัด internal ออกจากรายการแสดงผล
        tail = [c for c in tail if c not in internal_hide]
        df = _hide_empty_sp_group_ui(df.copy())
        df_ui = df.loc[:, ordered + tail]

        # ตั้งค่าตาราง
        self.result_table.setRowCount(len(df_ui))
        self.result_table.setColumnCount(len(df_ui.columns))
        self.result_table.setHorizontalHeaderLabels(df_ui.columns.tolist())

        # ให้หัวคอลัมน์เป็นเทาอ่อน และคงเส้นแบ่งคอลัมน์ไว้
        self.result_table.setShowGrid(True)
        self.result_table.horizontalHeader().setStyleSheet("""
        QHeaderView::section {
            background-color: #DCDCDC;  
            color: #1F2937;             
            font-weight: 600;
            padding: 6px 8px;
            border-top: 1px solid #BFBFBF;
            border-bottom: 1px solid #BFBFBF;
            border-right: 1px solid #BFBFBF;
            border-left: 0px;
        }
        QHeaderView::section:first {
            border-left: 1px solid #BFBFBF;
        }
        """)

        # Header "Verification" ให้เด่น
        if "Verification" in df_ui.columns:
            vcol = df_ui.columns.get_loc("Verification")
            head_item = QtWidgets.QTableWidgetItem("Verification")
            head_item.setForeground(QColor("black"))
            head_item.setToolTip("Auto/Manual verification status")
            self.result_table.setHorizontalHeaderItem(vcol, head_item)
            self.result_table.setColumnWidth(vcol, 120)

        symbol_names = {"Symbol/  Exact wording", "Symbol/ Exact wording", "Symbol/Exact wording"}
        sym_col = None
        for c in range(self.result_table.columnCount()):
            hi = self.result_table.horizontalHeaderItem(c)
            if hi and hi.text().strip() in symbol_names:
                sym_col = c
                break

        if sym_col is not None:
            hh = self.result_table.horizontalHeader()
            hh.setSectionResizeMode(sym_col, QHeaderView.Fixed) 
            self.result_table.setColumnWidth(sym_col, 320) 

        # Bold header font
        header_font = self.result_table.horizontalHeader().font()
        header_font.setBold(True)
        self.result_table.horizontalHeader().setFont(header_font)

        if hasattr(self, "update_row_highlight"):
            self.update_row_highlight()

        # ตั้งความกว้างหลัก
        equal_width = 240
        if "Requirement" in df_ui.columns:
            self.result_table.setColumnWidth(df_ui.columns.get_loc("Requirement"), equal_width)
        if "Specification" in df_ui.columns:
            self.result_table.setColumnWidth(df_ui.columns.get_loc("Specification"), equal_width)
        if "Symbol/ Exact wording" in df_ui.columns:
            self.result_table.setColumnWidth(df_ui.columns.get_loc("Symbol/ Exact wording"), 350)
        if "Package Panel" in df_ui.columns:
            self.result_table.setColumnWidth(df_ui.columns.get_loc("Package Panel"), equal_width)
        if "Procedure" in df_ui.columns:
            self.result_table.setColumnWidth(df_ui.columns.get_loc("Procedure"), equal_width)
        if "Remark" in df_ui.columns:
            self.result_table.setColumnWidth(df_ui.columns.get_loc("Remark"), 340)

        header = self.result_table.horizontalHeader()
        for i in range(df_ui.shape[1]):
            header.setSectionResizeMode(i, QtWidgets.QHeaderView.Fixed)

        try:
            if "Pages" in df_ui.columns:
                page_index = df_ui.columns.get_loc("Pages")
                ref_width = self.result_table.columnWidth(page_index)
            else:
                ref_width = 120
        except ValueError:
            ref_width = 120

        for col in ["Found", "Match", "Font Size", "Note", "Verification"]:
            if col in df_ui.columns:
                self.result_table.setColumnWidth(df_ui.columns.get_loc(col), ref_width)

        if "Note" in df_ui.columns:
            self.result_table.setColumnWidth(df_ui.columns.get_loc("Note"), 250)

        self.result_table.resizeRowsToContents()

        # แต่งสีคอลัมน์ Verification ให้สแกนง่าย
        if "Verification" in df_ui.columns:
            vcol = df_ui.columns.get_loc("Verification")
            for r in range(self.result_table.rowCount()):
                it = self.result_table.item(r, vcol)
                if not it:
                    continue
                txt = (it.text() or "").strip().lower()
                if txt == "manual":
                    it.setBackground(QColor("#FFF1B0")) 
                    it.setForeground(QColor("#8A6D00"))
                    it.setText("Manual")
                else:
                    it.setBackground(QColor("#E6F4EA")) 
                    it.setForeground(QColor("#0E7C3F"))
                    it.setText("Verified")
                it.setTextAlignment(QtCore.Qt.AlignCenter)

        # autosize คอลัมน์ยาว
        sym_header = None
        for name in ("Symbol/  Exact wording", "Symbol/ Exact wording", "Symbol/Exact wording"):
            if name in df_ui.columns:
                sym_header = name
                break
        if sym_header:
            self._autosize_column_to_contents(sym_header, min_w=360, max_w=720)
        if "Remark" in df_ui.columns:
            self._autosize_column_to_contents("Remark", min_w=340, max_w=560)

        # คำนวณความสูงใหม่หลังคอลัมน์กว้างขึ้น
        self.result_table.resizeRowsToContents()

        # helpers
        def linkify(text: str) -> str:
            if not isinstance(text, str) or not text:
                return "-"
            return re.sub(r'(https?://[^\s]+)', r'<a href="\1">\1</a>', text)

        def _lookup_image_groups(requirement_text: str, term_text: str):
            if getattr(self, "checklist_df", None) is None:
                return []
            req_series  = self.checklist_df.get("Requirement")
            term_series = self.checklist_df.get("Symbol/Exact wording")
            if req_series is None:
                return []

            req_left = req_series.astype(str).str.strip().str.lower()
            groups = []

            if term_series is not None:
                term_left = term_series.astype(str).str.strip().str.lower()
                if term_text and term_text.strip() and term_text.strip() != "-":
                    mask_rt = (req_left == str(requirement_text).strip().lower()) & \
                            (term_left == str(term_text).strip().lower())
                    sub = self.checklist_df[mask_rt]
                    if not sub.empty:
                        g = sub.iloc[0].get("Image_Groups_Resolved") or sub.iloc[0].get("Image_Groups", [])
                        groups = g or []

            if not groups:
                mask_r = (req_left == str(requirement_text).strip().lower())
                sub_r = self.checklist_df[mask_r]
                if not sub_r.empty:
                    for _, r in sub_r.iterrows():
                        g = r.get("Image_Groups_Resolved") or r.get("Image_Groups", [])
                        if g:
                            groups = g
                            break
            return groups or []

        # แคชภาพ
        if not hasattr(self, "_image_cache"):
            self._image_cache = {}

        for row_idx in range(len(df_ui)):
            row_ui  = df_ui.iloc[row_idx]   
            row_src = df_src.iloc[row_idx]     

            found        = str(row_ui.get("Found", ""))
            match        = str(row_ui.get("Match", ""))
            font_size    = str(row_ui.get("Font Size", ""))
            note         = str(row_ui.get("Note", ""))
            verification = str(row_ui.get("Verification", "")).strip().lower()

            for col_idx, header in enumerate(df_ui.columns):
                value = row_ui.get(header, "-")

                # --- Verification → แสดง Verified/Reject + tooltip ---
                if header == "Verification":
                    found     = str(row_ui.get("Found", ""))
                    match     = str(row_ui.get("Match", ""))
                    font_size = str(row_ui.get("Font Size", ""))
                    raw_verif = (row_ui.get("Verification", "") or "").strip().lower()
                    is_manual = (raw_verif == "manual") 

                    if is_manual:
                        text = "Manual"
                        ok = None
                    else:
                        ok_found = found.strip().startswith("✅")
                        ok_match = match.strip().startswith("✔")
                        fs = font_size.strip()
                        ok_fsize = (fs in ("", "-")) or fs.startswith("✔")
                        ok = ok_found and ok_match and ok_fsize
                        text = "Verified" if ok else "Rejected"

                    item = QtWidgets.QTableWidgetItem(text)
                    item.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
                    item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)

                    f = item.font()
                    f.setBold(True)
                    item.setFont(f)

                    if not is_manual:
                        item.setForeground(QColor("#15803d") if ok else QColor("red"))
                    else:
                        item.setForeground(QColor("black"))

                    tip = f"Found: {found or '-'}\nMatch: {match or '-'}\nFont Size: {font_size or '-'}"
                    if is_manual:
                        tip += "\n— Manual check"

                    item.setToolTip(tip)
                    item.setData(QtCore.Qt.UserRole, "manual" if is_manual else "auto")

                    self.result_table.setItem(row_idx, col_idx, item)
                    continue

                if header == "Symbol/ Exact wording":
                    req_text  = str(row_src.get("Requirement", "")).strip()

                    def _clean_plain(s: str) -> str:
                        if s is None:
                            return ""
                        s = str(s).strip()
                        return "" if s.lower() in ("nan", "none", "-") else s

                    # กลุ่มรูปของแถวนี้
                    groups = row_src.get("Image_Groups_Resolved") or row_src.get("Image_Groups") or []
                    has_images = bool(groups and any(g.get("paths") for g in groups))

                    # plain text (จาก df_ui)
                    term_raw  = row_ui.get(header, "")
                    term_text = _clean_plain(term_raw)

                    # html (underline/bold) จาก excel
                    def _clean_html(s: str) -> str:
                        if not isinstance(s, str):
                            return ""
                        s2 = s.strip()
                        if s2.lower() in ("nan", "none", "-"):
                            return ""
                        plain = re.sub(r"<[^>]+>", "", s2).strip()
                        return "" if plain == "" else s2

                    html_val = ""
                    for k in ("__Term_HTML__", "Term_Underline_HTML"):
                        v = row_src.get(k, "")
                        if isinstance(v, str) and v.strip():
                            html_val = _clean_html(v)
                            if html_val:
                                break

                    # ตัดสินใจข้อความที่จะแสดง
                    def _norm_basic(s: str) -> str:
                        s = str(s or "").replace("\u00a0", " ")
                        s = _ud.normalize("NFKD", s)
                        s = "".join(ch for ch in s if not _ud.combining(ch))
                        return re.sub(r"\s+", " ", s).strip().lower()

                    def _sameish(a: str, b: str) -> bool:
                        A, B = _norm_basic(a), _norm_basic(b)
                        if not A or not B:
                            return False
                        if A == B or (A in B) or (B in A):
                            return True
                        ta = {t for t in A.split() if len(t) > 1}
                        tb = {t for t in B.split() if len(t) > 1}
                        if not ta or not tb:
                            return False
                        inter = len(ta & tb)
                        return inter / max(len(ta), len(tb)) >= 0.6 

                    html_plain = re.sub(r"<[^>]+>", "", html_val) if html_val else ""

                    if html_val:
                        display_text = html_val
                        plain_for_measure = html_plain or term_text
                    elif term_text:
                        display_text = term_text
                        plain_for_measure = term_text
                    else:
                        display_text = "" if has_images else "-"
                        plain_for_measure = ""

                    # สร้าง UI ของเซลล์
                    container = QtWidgets.QWidget()
                    outer = QtWidgets.QVBoxLayout(container)
                    outer.setContentsMargins(4, 2, 4, 2)
                    outer.setSpacing(4)
                    outer.setAlignment(QtCore.Qt.AlignCenter)

                    # วางข้อความเฉพาะเมื่อมีข้อความจริงเท่านั้น
                    term_label = None
                    if display_text.strip():
                        term_label = QtWidgets.QLabel()
                        term_label.setFont(self.result_table.font())  # ใช้ฟอนต์ของตารางเสมอ
                        term_label.setTextInteractionFlags(Qt.TextBrowserInteraction)
                        term_label.setAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)

                        is_html_mode = bool(html_val)
                        plain_for_measure = re.sub(r"<[^>]+>", "", display_text or "")

                        if is_html_mode:
                            inner_html = re.sub(r"\r\n|\r|\n", "<br/>", display_text or "")
                        else:
                            inner_html = _html.escape((display_text or "").replace("\r", ""))
                            inner_html = inner_html.replace("\n", "<br/>")

                        term_label.setProperty("base_inner_html", inner_html)

                        term_label.setTextFormat(QtCore.Qt.RichText)
                        term_label.setText(self._wrap_html_with_table_font(inner_html))

                        # ความกว้างวัดการตัดบรรทัดคงเดิม
                        col_width = self.result_table.columnWidth(col_idx)
                        fm = term_label.fontMetrics()
                        inner_w = max(40, col_width - 12)
                        term_label.setMinimumWidth(inner_w)
                        term_label.setMaximumWidth(inner_w)
                        term_label.setWordWrap(fm.horizontalAdvance(plain_for_measure) > inner_w if plain_for_measure else True)

                        # คงสีแดงกรณี Not Found
                        if str(row_ui.get("Found", "")).startswith("❌"):
                            term_label.setStyleSheet(term_label.styleSheet() + f" color:{RED_HEX};")

                        outer.addWidget(term_label, 0, QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)

                    # ส่วนรูปภาพ
                    if has_images:
                        all_paths = []
                        for g in groups:
                            all_paths.extend(g.get("paths", []))

                        if not hasattr(self, "_image_cache"):
                            self._image_cache = {}

                        if all_paths:
                            img_wrap = QtWidgets.QWidget()
                            img_vbox = QtWidgets.QVBoxLayout(img_wrap)
                            img_vbox.setContentsMargins(IMG_SIDE_PADDING, 0, IMG_SIDE_PADDING, 0)
                            img_vbox.setSpacing(8)
                            img_vbox.setAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)

                            # ความกว้างรูปสูงสุด = ความกว้างคอลัมน์ - padding ซ้าย/ขวา
                            col_width = self.result_table.columnWidth(col_idx)
                            max_img_w = max(40, col_width - 2 * IMG_SIDE_PADDING)

                            for p in all_paths:
                                if not p:
                                    continue

                                pm = self._image_cache.get(p)
                                if pm is None:
                                    qpm = QtGui.QPixmap(p)
                                    pm = qpm if not qpm.isNull() else None
                                    self._image_cache[p] = pm if pm else QtGui.QPixmap()

                                lbl = QtWidgets.QLabel()
                                lbl.setAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)

                                # เก็บเมตาไว้ใช้ตอนรีสเกลเมื่อคอลัมน์ถูกปรับ
                                lbl.setProperty("img_path", p)
                                is_logo = _is_logo_name(p, req_text) and not _must_fill_width(req_text)
                                lbl.setProperty("is_logo", is_logo)

                                if not pm:
                                    lbl.setText(f"[!] Missing image: {p}")
                                    lbl.setStyleSheet(f"color:{RED_HEX};")
                                    img_vbox.addWidget(lbl, 0, QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
                                    continue

                                col_width = self.result_table.columnWidth(col_idx)
                                max_img_w = max(40, col_width - 2 * IMG_SIDE_PADDING)

                                # รูปทั่วไปขยายเกือบเต็มคอลัมน์, โลโก้/มาร์ก: จำกัดไม่ให้ใหญ่เกิน
                                target_w = min(LOGO_MAX_WIDTH_PX, max_img_w) if is_logo else int(max_img_w * 0.98)
                                scaled = pm.scaledToWidth(max(1, target_w), QtCore.Qt.SmoothTransformation)
                                lbl.setPixmap(scaled)

                                img_vbox.addWidget(lbl, 0, QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)

                            if not display_text.strip():
                                outer.addStretch(1)
                                outer.addWidget(img_wrap, 0, QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
                                outer.addStretch(1)
                            else:
                                outer.addWidget(img_wrap, 0, QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)

                    container.setLayout(outer)
                    self.result_table.takeItem(row_idx, col_idx)
                    self.result_table.setCellWidget(row_idx, col_idx, container)
                    self._attach_row_select(container, row_idx)

                    # ปรับความสูงแถวให้พอดีเนื้อหา
                    self.result_table.resizeRowToContents(row_idx)
                    if self.result_table.rowHeight(row_idx) < 28:
                        self.result_table.setRowHeight(row_idx, 28)
                    continue

                # Remark: จัดกึ่งกลางเสมอ + ตัดบรรทัดอัตโนมัติ
                if header == "Remark":
                    URL_RX = re.compile(r'(https?://[^\s<>"\')]+|www\.[^\s<>"\')]+)', re.IGNORECASE)

                    url_from_col = str(row_src.get("Remark URL", "") or row_src.get("Remark Link", "") or "").strip()
                    txt = (str(value) if value is not None else "").strip()

                    # เคสเป็น "-" หรือว่าง และไม่มี URL → แสดง "-" กลางเซลล์
                    if (txt in ("", "-", "–", "—", "=")) and (not url_from_col):
                        item = QtWidgets.QTableWidgetItem("-")
                        item.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
                        item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
                        self.result_table.setItem(row_idx, col_idx, item)
                        continue

                    # container + layout
                    rwrap = QtWidgets.QWidget()
                    rlay = QtWidgets.QVBoxLayout(rwrap)
                    rlay.setContentsMargins(6, 2, 6, 2)
                    rlay.setSpacing(0)

                    lbl = QtWidgets.QLabel()
                    lbl.setTextFormat(QtCore.Qt.RichText)
                    lbl.setOpenExternalLinks(True)
                    lbl.setTextInteractionFlags(QtCore.Qt.TextBrowserInteraction)
                    lbl.setAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
                    lbl.setWordWrap(True)
                    lbl.setFont(self.result_table.font())
                    lbl.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
                    lbl.setStyleSheet("QLabel { padding: 0; margin: 0; }")

                    # คำนวณความกว้างภายใน
                    col_w   = self.result_table.columnWidth(col_idx)
                    inner_w = max(40, col_w - 12)
                    fm = lbl.fontMetrics()

                    def _linkify_plain_to_html(s: str) -> str:
                        if not s:
                            return "-"
                        s = s.replace("\r\n", "\n").replace("\r", "\n")
                        parts, last = [], 0
                        for m in URL_RX.finditer(s):
                            parts.append(html.escape(s[last:m.start()]))
                            raw = m.group(1)
                            href = raw if raw.lower().startswith(("http://","https://")) else ("http://" + raw)
                            parts.append(f'<a href="{html.escape(href)}">{html.escape(raw)}</a>')
                            last = m.end()
                        parts.append(html.escape(s[last:]))
                        return "<br>".join(p or "" for p in "".join(parts).split("\n")) or "-"

                    has_url_in_text = bool(URL_RX.search(txt)) if txt not in ("", "-", "–", "—") else False

                    if has_url_in_text:
                        content_html = _linkify_plain_to_html(txt)
                    else:
                        content_html = self._remark_pairs_to_html(txt, inner_w, fm)

                    if url_from_col and not has_url_in_text and content_html.strip() != "-":
                        content_html = self._wrap_all_as_link(content_html, url_from_col)

                    lbl.setText(self._wrap_html_with_table_font(content_html))
                    lbl.setMinimumWidth(inner_w)
                    lbl.setMaximumWidth(inner_w)

                    lbl.setProperty("raw_remark", txt)
                    lbl.setProperty("has_url_in_text", has_url_in_text)
                    lbl.setProperty("remark_url_from_col", url_from_col)

                    # ควบคุม word-wrap ตามความกว้างจริง
                    plain = re.sub(r"<[^>]+>", "", content_html or "")
                    lbl.setWordWrap(fm.horizontalAdvance(plain) > inner_w if plain else True)

                    rlay.addWidget(lbl)
                    self.result_table.setCellWidget(row_idx, col_idx, rwrap)
                    self._attach_row_select(rwrap, row_idx)
                    continue

                # คอลัมน์อื่นๆ 
                item = QtWidgets.QTableWidgetItem(str(value))
                item.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
                item.setToolTip(str(value))
                item.setText(str(value))

                # Alignment: Requirement = ซ้าย/หนา, อื่นๆ = กึ่งกลาง
                if header == "Requirement":
                    item.setTextAlignment(QtCore.Qt.AlignCenter | QtCore.Qt.AlignVCenter)
                    f = item.font(); f.setBold(True); item.setFont(f)
                else:
                    item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
                
                # สีตามสถานะ
                if verification == "manual":
                    item.setBackground(QColor("#fff9cc"))
                    if header in ["Found", "Match", "Font Size", "Note"]:
                        item.setForeground(QColor("gray"))
                else:
                    if found.startswith("❌") and header != "Requirement":
                        item.setForeground(QColor(RED_HEX))

                    elif header == "Match" and match.startswith("❌"):
                        item.setForeground(QColor("red"))
                    elif header == "Font Size":
                        if found.startswith("❌") and not font_size.startswith("✔"):
                            item.setForeground(QColor("red"))
                    elif header == "Note" and note.strip() not in ["-", ""]:
                        item.setForeground(QColor("red"))

                self.result_table.setItem(row_idx, col_idx, item)
            self.result_table.resizeRowsToContents()

    def _on_column_resized(self, logicalIndex: int, oldSize: int, newSize: int):
        header_item = self.result_table.horizontalHeaderItem(logicalIndex)
        if not header_item:
            return
        name = header_item.text().strip().lower()
        if name not in ("symbol/ exact wording", "symbol/  exact wording", "symbol/exact wording", "remark"):
            return

        col_idx = logicalIndex
        col_w = self.result_table.columnWidth(col_idx)
        inner_w = max(40, col_w - 12)

        for row in range(self.result_table.rowCount()):
            cell = self.result_table.cellWidget(row, col_idx)
            if not cell:
                continue

            labels = cell.findChildren(QtWidgets.QLabel)
            for lbl in labels:
                html = lbl.text() if hasattr(lbl, "text") else ""
                if not isinstance(html, str):
                    continue

                # รีสเกลข้อความทั้ง Symbol และ Remark
                txt_plain = re.sub(r"<[^>]+>", "", html)
                fm = lbl.fontMetrics()
                lbl.setMinimumWidth(inner_w)
                lbl.setMaximumWidth(inner_w)
                lbl.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
                need_wrap = fm.horizontalAdvance(txt_plain) > inner_w if txt_plain else True
                lbl.setWordWrap(need_wrap)

                raw = lbl.property("raw_remark")
                if raw is not None and name == "remark":
                    has_url = bool(lbl.property("has_url_in_text"))
                    url_from_col = lbl.property("remark_url_from_col") or ""
                    if not has_url:
                        new_html = self._remark_pairs_to_html(str(raw), inner_w, fm)
                        if url_from_col:
                            new_html = self._wrap_all_as_link(new_html, url_from_col)
                        lbl.setText(self._wrap_html_with_table_font(new_html))

            # รีสเกลรูป เฉพาะคอลัมน์ Symbol
            for img_lbl in labels:
                img_path = img_lbl.property("img_path")
                if not img_path:
                    continue
                pm_orig = getattr(self, "_image_cache", {}).get(img_path)
                if not isinstance(pm_orig, QtGui.QPixmap) or pm_orig.isNull():
                    continue
                max_img_w = max(40, col_w - 2 * IMG_SIDE_PADDING)
                is_logo   = bool(img_lbl.property("is_logo"))
                target_w  = min(LOGO_MAX_WIDTH_PX, max_img_w) if is_logo else int(max_img_w * 0.98)
                img_lbl.setPixmap(pm_orig.scaledToWidth(max(1, target_w), QtCore.Qt.SmoothTransformation))

    def export_results(self):
        if self.result_df is None or self.result_df.empty:
            QtWidgets.QMessageBox.warning(self, "No Results", "Please run checking before exporting.")
            return
        export_result_to_excel(self.result_df)

    def preview_pdf(self):
        if not getattr(self, "pdf_path", None):
            QtWidgets.QMessageBox.information(self, "Preview PDF", "กรุณาอัปโหลดไฟล์ PDF ก่อน")
            return
        if getattr(self, "checklist_df", None) is None:
            QtWidgets.QMessageBox.information(self, "Preview PDF", "กรุณาอัปโหลดไฟล์ Checklist ก่อน")
            return

        df = getattr(self, "result_df", None)
        if df is None or df.empty:
            QtWidgets.QMessageBox.information(self, "Preview PDF", "ยังไม่มีผลตรวจสำหรับพรีวิว กรุณากด Start Checking ก่อน")
            return

        rows = []
        cols = df.columns.str.strip().tolist()
        col_req   = "Requirement" if "Requirement" in cols else None
        col_sym   = "Symbol/ Exact wording" if "Symbol/ Exact wording" in cols else None
        col_found = "Found" if "Found" in cols else None
        col_ver   = "Verification" if "Verification" in cols else None
        col_pages = "Pages" if "Pages" in cols else None

        if not (col_sym and col_found):
            QtWidgets.QMessageBox.information(self, "Preview PDF", "ไม่พบคอลัมน์ที่จำเป็น (Symbol/Exact wording, Found)")
            return

        for i, row in df.iterrows():
            symbol = str(row.get(col_sym, "") or "").strip()
            req    = str(row.get(col_req, "") or "").strip() if col_req else f"Row {i+1}"
            found  = str(row.get(col_found, "") or "")
            ver    = str(row.get(col_ver, "") or "").strip().lower() if col_ver else ""
            pages  = str(row.get(col_pages, "") or "") if col_pages else ""

            # --- สถานะ + หน้าที่ใช้จริงในพรีวิว ---
            is_found = found.strip().startswith("✅")
            is_manual = (ver == "manual")
            status = "manual" if is_manual else ("found" if is_found else "missing")

            # ถ้าไม่พบ → pages_spec = "" (จะถูก parse เป็น set() = no pages)
            pages_spec = pages if ((is_found or is_manual) and pages not in ("", "-", "—")) else ""
            
            try:
                row_id = int(i)
            except Exception:
                row_id = len(rows)
            
            rows.append({
                "id": row_id,
                "requirement": req,
                "symbol": symbol,
                "status": status,
                "pages_spec": pages_spec,
            })

        # ==== PRUNE rows: อย่าให้ SPW (สั้น) ที่ Not Found หลุดไปไฮไลท์ เมื่อมี SPG (Found) อยู่แล้ว ====
        try:
            req_col   = col_req   or "Requirement"
            pages_col = col_pages or "Pages"
            have_spg_found = False
            if req_col in df.columns and pages_col in df.columns:
                def _pages_not_empty(v):
                    if v is None: return False
                    if isinstance(v, (set, list, tuple)): return len(v) > 0
                    s = str(v).strip()
                    return not (s == "" or s == "-" or s == "—" or s.lower() == "none" or s == "0")
                reqn = df[req_col].fillna("").str.lower()
                m_spg = reqn.str.contains("international warning statement", regex=False) & reqn.str.contains(r"\bspg\b", regex=True)
                if m_spg.any():
                    have_spg_found = bool(df.loc[m_spg, pages_col].map(_pages_not_empty).any())

            if have_spg_found:

                def _is_short_spw_symbol(s: str) -> bool:
                    s2 = (s or "").lower()
                    return ("warning" in s2 and "small parts" in s2 and "may be generat" not in s2)

                rows = [r for r in rows if not (_is_short_spw_symbol(r.get("symbol")) and r.get("status") != "found")]
        except Exception:
            pass

        # ถ้ามีหน้าต่างพรีวิวเปิดอยู่แล้ว ให้โฟกัสแทนการเปิดใหม่
        w = getattr(self, "_pdf_preview_win", None)
        if isinstance(w, QtWidgets.QWidget) and w.isVisible():
            w.activateWindow()
            w.raise_()
            return

        # ถ้ามีอ้างอิงหน้าต่างเดิม ให้ปิดและเคลียร์ก่อน
        if isinstance(w, QtWidgets.QWidget):
            try:
                w.close()
            except Exception:
                pass
        self._pdf_preview_win = None

        # เปิดหน้าต่างพรีวิวแบบ top-level ที่ย่อ/ขยายได้
        self._pdf_preview_win = PdfPreviewWindow(pdf_path=self.pdf_path, rows=rows, parent=None)
        self._pdf_preview_win.destroyed.connect(lambda: setattr(self, "_pdf_preview_win", None))
        self._pdf_preview_win.show()
        self._pdf_preview_win.activateWindow()
        self._pdf_preview_win.raise_()

    def _on_table_cell_clicked(self, row: int, col: int):
        self._set_hide_selection_overlay_during_search(False)

    def _on_table_cell_double_clicked(self, row: int, col: int):
        sel = self.result_table.selectionModel()
        if sel and sel.isRowSelected(row, self.result_table.rootIndex()):
            self.result_table.clearSelection()

    def search_text(self):
        query = (self.search_input.text() or "").strip()
        if not query:
            QtWidgets.QMessageBox.information(self, "Search", "Please enter a term to search.")
            return
        if self.result_table.rowCount() == 0:
            QtWidgets.QMessageBox.information(self, "Search", "No results to search.")
            return

        first_row = self._apply_symbol_highlight(query)
        if first_row >= 0:
            self._set_hide_selection_overlay_during_search(False)

            sym_idx = self.get_column_index("Symbol/ Exact wording")
            if sym_idx < 0:
                sym_idx = 0
            index = self.result_table.model().index(first_row, sym_idx)
            self.result_table.scrollTo(index, QtWidgets.QAbstractItemView.PositionAtCenter)

            self._select_entire_row(first_row)
            self.update_row_highlight()
            return

        QtWidgets.QMessageBox.information(self, "Search", f"'{query}' not found in Symbol column.")
        self._set_hide_selection_overlay_during_search(False)