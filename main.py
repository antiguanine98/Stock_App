#!/usr/bin/env python3
"""재고 소급 보정 · 추이 차트 · AI 분석 애플리케이션.

엑셀 양식(테스트파일1.xlsx):
  - 기본정보: 순번, 표준품구분, 관리번호, 한글명, 영문명, 잔고, 등록일자, 분양여부
  - 변동기록: 변경일자N / 재고량N 쌍 (I열 이후)
  - 식별자: 관리번호 + 한글명
"""

from __future__ import annotations

import os
import sys
from typing import Optional

import matplotlib
import pandas as pd
from matplotlib import font_manager
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QAction, QDragEnterEvent, QDropEvent, QFont
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QSplitter,
    QStatusBar,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from stock_logic import (
    StockItem,
    _cell_str,
    _is_empty,
    build_ai_prompt,
    format_date,
    items_for_ai_analysis,
    load_stock_excel,
)

# ---------------------------------------------------------------------------
# 테마 (딥 블루 / 네이비 + 연한 회색·화이트)
# ---------------------------------------------------------------------------

COLORS = {
    "navy": "#0F2744",
    "navy_mid": "#1A3A5C",
    "blue": "#2B6CB0",
    "blue_soft": "#3B82C4",
    "accent": "#2563EB",
    "bg": "#F4F6F9",
    "surface": "#FFFFFF",
    "border": "#D7DEE8",
    "border_soft": "#E6EBF2",
    "text": "#1E293B",
    "muted": "#64748B",
    "success": "#059669",
    "warn": "#D97706",
    "danger": "#DC2626",
    "table_alt": "#F8FAFC",
    "table_header": "#0F2744",
    "chart": "#1A3A5C",
}


APP_STYLESHEET = f"""
QMainWindow, QWidget {{
    background-color: {COLORS['bg']};
    color: {COLORS['text']};
    font-family: "Segoe UI", "Malgun Gothic", "Apple SD Gothic Neo",
                 "Noto Sans CJK KR", "WenQuanYi Micro Hei", sans-serif;
    font-size: 13px;
}}
QMenuBar {{
    background: {COLORS['surface']};
    border-bottom: 1px solid {COLORS['border']};
    padding: 2px 6px;
}}
QMenuBar::item:selected {{
    background: {COLORS['border_soft']};
    border-radius: 4px;
}}
QLabel#AppTitle {{
    color: {COLORS['navy']};
    font-size: 20px;
    font-weight: 700;
}}
QLabel#AppSubtitle {{
    color: {COLORS['muted']};
    font-size: 12px;
}}
QLabel#SectionHint {{
    color: {COLORS['muted']};
    font-size: 12px;
}}
QLabel#FieldLabel {{
    color: {COLORS['navy_mid']};
    font-weight: 600;
}}
QFrame#Card {{
    background: {COLORS['surface']};
    border: 1px solid {COLORS['border']};
    border-radius: 12px;
}}
QFrame#DropZone {{
    background: {COLORS['surface']};
    border: 2px dashed {COLORS['blue_soft']};
    border-radius: 12px;
}}
QFrame#DropZone[dragActive="true"] {{
    background: #EEF5FC;
    border: 2px dashed {COLORS['accent']};
}}
QLabel#DropTitle {{
    color: {COLORS['navy']};
    font-size: 15px;
    font-weight: 700;
}}
QLabel#DropHint {{
    color: {COLORS['muted']};
    font-size: 12px;
}}
QLineEdit, QComboBox, QTextEdit {{
    background: {COLORS['surface']};
    border: 1px solid {COLORS['border']};
    border-radius: 8px;
    padding: 8px 12px;
    selection-background-color: {COLORS['blue']};
}}
QLineEdit:focus, QComboBox:focus, QTextEdit:focus {{
    border: 1px solid {COLORS['accent']};
}}
QComboBox::drop-down {{
    border: none;
    width: 28px;
}}
QPushButton {{
    background: {COLORS['navy']};
    color: white;
    border: none;
    border-radius: 8px;
    padding: 9px 16px;
    font-weight: 600;
}}
QPushButton:hover {{
    background: {COLORS['navy_mid']};
}}
QPushButton:pressed {{
    background: {COLORS['blue']};
}}
QPushButton#SecondaryButton {{
    background: {COLORS['surface']};
    color: {COLORS['navy']};
    border: 1px solid {COLORS['border']};
}}
QPushButton#SecondaryButton:hover {{
    background: {COLORS['border_soft']};
}}
QPushButton#AccentButton {{
    background: {COLORS['accent']};
}}
QPushButton#AccentButton:hover {{
    background: {COLORS['blue']};
}}
QTabWidget::pane {{
    background: {COLORS['surface']};
    border: 1px solid {COLORS['border']};
    border-radius: 12px;
    top: -1px;
    padding: 8px;
}}
QTabBar::tab {{
    background: transparent;
    color: {COLORS['muted']};
    padding: 10px 18px;
    margin-right: 4px;
    border-top-left-radius: 8px;
    border-top-right-radius: 8px;
    font-weight: 600;
}}
QTabBar::tab:selected {{
    background: {COLORS['surface']};
    color: {COLORS['navy']};
    border: 1px solid {COLORS['border']};
    border-bottom: 2px solid {COLORS['accent']};
}}
QTabBar::tab:hover:!selected {{
    color: {COLORS['navy_mid']};
    background: {COLORS['border_soft']};
}}
QTableWidget {{
    background: {COLORS['surface']};
    alternate-background-color: {COLORS['table_alt']};
    gridline-color: {COLORS['border_soft']};
    border: 1px solid {COLORS['border_soft']};
    border-radius: 8px;
    selection-background-color: #DBEAFE;
    selection-color: {COLORS['navy']};
}}
QHeaderView::section {{
    background: {COLORS['table_header']};
    color: white;
    padding: 8px;
    border: none;
    font-weight: 600;
}}
QStatusBar {{
    background: {COLORS['surface']};
    border-top: 1px solid {COLORS['border']};
}}
QLabel#StatusChip {{
    color: {COLORS['text']};
    padding: 2px 10px;
}}
QLabel#StatusDivider {{
    color: {COLORS['border']};
}}
QScrollBar:vertical {{
    background: transparent;
    width: 10px;
    margin: 2px;
}}
QScrollBar::handle:vertical {{
    background: #C5D0DE;
    border-radius: 5px;
    min-height: 24px;
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0;
}}
"""


def _configure_matplotlib_font() -> None:
    """한글 축/제목 표시를 위한 폰트 설정."""
    candidates = [
        "WenQuanYi Micro Hei",
        "Noto Sans CJK KR",
        "NanumGothic",
        "Malgun Gothic",
        "AppleGothic",
    ]
    available = {f.name for f in font_manager.fontManager.ttflist}
    for name in candidates:
        if name in available:
            matplotlib.rcParams["font.family"] = name
            break
    matplotlib.rcParams["axes.unicode_minus"] = False


_configure_matplotlib_font()


class GeminiWorker(QThread):
    finished_ok = Signal(str)
    failed = Signal(str)

    def __init__(self, prompt: str, api_key: str, parent=None):
        super().__init__(parent)
        self.prompt = prompt
        self.api_key = api_key

    def run(self) -> None:
        try:
            import google.generativeai as genai

            genai.configure(api_key=self.api_key)
            model_name = os.getenv("GEMINI_MODEL", "gemini-1.5-flash")
            model = genai.GenerativeModel(model_name)
            response = model.generate_content(self.prompt)
            text = getattr(response, "text", None) or str(response)
            self.finished_ok.emit(text.strip())
        except Exception as exc:  # noqa: BLE001 - UI 전달용
            self.failed.emit(str(exc))


class MplCanvas(FigureCanvas):
    def __init__(self, parent=None):
        self.figure = Figure(figsize=(8, 4.5), facecolor=COLORS["surface"], tight_layout=True)
        super().__init__(self.figure)
        self.setParent(parent)
        self.ax = self.figure.add_subplot(111)
        self.ax.set_facecolor(COLORS["surface"])


class DropZone(QFrame):
    """점선 스타일 엑셀 드래그앤드롭 패널."""

    file_dropped = Signal(str)
    browse_clicked = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("DropZone")
        self.setAcceptDrops(True)
        self.setMinimumHeight(118)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.setProperty("dragActive", False)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 16, 20, 16)
        lay.setSpacing(6)
        lay.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.title = QLabel("엑셀 파일을 여기에 끌어다 놓으세요")
        self.title.setObjectName("DropTitle")
        self.title.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.hint = QLabel("또는 클릭하여 파일 선택  ·  .xlsx / .xls")
        self.hint.setObjectName("DropHint")
        self.hint.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.file_label = QLabel("선택된 파일 없음")
        self.file_label.setObjectName("DropHint")
        self.file_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.file_label.setWordWrap(True)

        btn_row = QHBoxLayout()
        btn_row.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.browse_btn = QPushButton("파일 찾아보기")
        self.browse_btn.setObjectName("SecondaryButton")
        self.browse_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.browse_btn.clicked.connect(self.browse_clicked.emit)
        btn_row.addWidget(self.browse_btn)

        lay.addWidget(self.title)
        lay.addWidget(self.hint)
        lay.addLayout(btn_row)
        lay.addWidget(self.file_label)

    def set_file_name(self, path: str) -> None:
        name = os.path.basename(path) if path else "선택된 파일 없음"
        self.file_label.setText(name)
        self.title.setText("파일이 준비되었습니다" if path else "엑셀 파일을 여기에 끌어다 놓으세요")

    def _set_drag_active(self, active: bool) -> None:
        self.setProperty("dragActive", active)
        self.style().unpolish(self)
        self.style().polish(self)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if url.toLocalFile().lower().endswith((".xlsx", ".xls")):
                    event.acceptProposedAction()
                    self._set_drag_active(True)
                    return
        event.ignore()

    def dragLeaveEvent(self, event) -> None:  # noqa: ANN001
        self._set_drag_active(False)
        super().dragLeaveEvent(event)

    def dropEvent(self, event: QDropEvent) -> None:
        self._set_drag_active(False)
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith((".xlsx", ".xls")):
                self.file_dropped.emit(path)
                event.acceptProposedAction()
                return
        event.ignore()

    def mousePressEvent(self, event) -> None:  # noqa: ANN001
        if event.button() == Qt.MouseButton.LeftButton:
            # 버튼 영역이 아니면 찾아보기로 연결
            child = self.childAt(event.position().toPoint())
            if child is not self.browse_btn:
                self.browse_clicked.emit()
        super().mousePressEvent(event)


class StockApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Stock Insight — 재고 소급 보정 · 추이 차트 · AI 분석")
        self.resize(1320, 860)

        self.items: list[StockItem] = []
        self.corrected_df: Optional[pd.DataFrame] = None
        self._item_by_manage_no: dict[str, StockItem] = {}
        self._gemini_worker: Optional[GeminiWorker] = None
        self._loaded_path: str = ""
        self._correction_count: int = 0
        self._all_combo_labels: list[tuple[str, str]] = []

        self._build_status_bar()
        self._build_ui()
        self._build_menu()
        self.refresh_api_status()

    def _build_menu(self) -> None:
        open_action = QAction("엑셀 열기...", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.open_excel)
        file_menu = self.menuBar().addMenu("파일")
        file_menu.addAction(open_action)

    def _build_status_bar(self) -> None:
        bar = QStatusBar()
        bar.setSizeGripEnabled(False)
        self.setStatusBar(bar)

        self.status_file = QLabel("파일: 없음")
        self.status_file.setObjectName("StatusChip")
        self.status_correction = QLabel("소급 보정: 0건")
        self.status_correction.setObjectName("StatusChip")
        self.status_api = QLabel("API: 미연결")
        self.status_api.setObjectName("StatusChip")

        def divider() -> QLabel:
            d = QLabel("│")
            d.setObjectName("StatusDivider")
            return d

        bar.addWidget(self.status_file, 1)
        bar.addWidget(divider())
        bar.addWidget(self.status_correction, 1)
        bar.addWidget(divider())
        bar.addWidget(self.status_api, 1)

    def _build_ui(self) -> None:
        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)
        layout.setContentsMargins(18, 14, 18, 10)
        layout.setSpacing(12)

        # --- 헤더 ---
        header = QVBoxLayout()
        header.setSpacing(2)
        title = QLabel("Stock Insight")
        title.setObjectName("AppTitle")
        subtitle = QLabel("소급 보정된 재고 데이터를 표·차트·AI 리포트로 한눈에 확인합니다")
        subtitle.setObjectName("AppSubtitle")
        header.addWidget(title)
        header.addWidget(subtitle)
        layout.addLayout(header)

        # --- 상단: 드롭존 + API 키 ---
        top_card = QFrame()
        top_card.setObjectName("Card")
        top_lay = QVBoxLayout(top_card)
        top_lay.setContentsMargins(14, 14, 14, 14)
        top_lay.setSpacing(12)

        self.drop_zone = DropZone()
        self.drop_zone.browse_clicked.connect(self.open_excel)
        self.drop_zone.file_dropped.connect(self.load_file)
        top_lay.addWidget(self.drop_zone)

        api_row = QHBoxLayout()
        api_row.setSpacing(10)
        api_label = QLabel("Gemini API Key")
        api_label.setObjectName("FieldLabel")
        self.api_key_edit = QLineEdit()
        self.api_key_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.api_key_edit.setPlaceholderText("Gemini API Key를 입력하세요")
        env_key = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY") or ""
        if env_key:
            self.api_key_edit.setText(env_key)
        self.api_key_edit.textChanged.connect(self.refresh_api_status)
        self.api_toggle_btn = QPushButton("표시")
        self.api_toggle_btn.setObjectName("SecondaryButton")
        self.api_toggle_btn.setFixedWidth(72)
        self.api_toggle_btn.clicked.connect(self.toggle_api_visibility)
        api_row.addWidget(api_label)
        api_row.addWidget(self.api_key_edit, stretch=1)
        api_row.addWidget(self.api_toggle_btn)
        top_lay.addLayout(api_row)
        layout.addWidget(top_card)

        # --- 탭 ---
        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        layout.addWidget(self.tabs, stretch=1)

        # 보정 데이터 표
        table_page = QWidget()
        table_layout = QVBoxLayout(table_page)
        table_layout.setContentsMargins(8, 8, 8, 8)
        table_layout.setSpacing(8)
        hint = QLabel("행을 더블클릭하면 해당 관리번호의 [재고 추이 차트]로 바로 이동합니다.")
        hint.setObjectName("SectionHint")
        table_layout.addWidget(hint)
        self.table = QTableWidget()
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.setShowGrid(False)
        self.table.verticalHeader().setVisible(False)
        self.table.setSortingEnabled(True)
        self.table.cellDoubleClicked.connect(self.on_table_double_clicked)
        table_layout.addWidget(self.table)
        self.tabs.addTab(table_page, "보정 데이터 표")

        # 재고 추이 차트
        chart_page = QWidget()
        chart_layout = QVBoxLayout(chart_page)
        chart_layout.setContentsMargins(8, 8, 8, 8)
        chart_layout.setSpacing(10)

        selector_card = QFrame()
        selector_card.setObjectName("Card")
        selector_lay = QHBoxLayout(selector_card)
        selector_lay.setContentsMargins(12, 10, 12, 10)
        selector_lay.setSpacing(10)
        selector_label = QLabel("품목 검색 / 선택")
        selector_label.setObjectName("FieldLabel")
        self.item_combo = QComboBox()
        self.item_combo.setEditable(True)
        self.item_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.item_combo.setPlaceholderText("관리번호(한글명)로 검색하세요")
        self.item_combo.currentIndexChanged.connect(self.on_combo_changed)
        self.item_combo.lineEdit().textEdited.connect(self.filter_combo_items)
        selector_lay.addWidget(selector_label)
        selector_lay.addWidget(self.item_combo, stretch=1)
        chart_layout.addWidget(selector_card)

        chart_frame = QFrame()
        chart_frame.setObjectName("Card")
        chart_frame_lay = QVBoxLayout(chart_frame)
        chart_frame_lay.setContentsMargins(10, 10, 10, 10)
        self.canvas = MplCanvas(self)
        chart_frame_lay.addWidget(self.canvas)
        self.chart_info = QLabel("품목을 선택하면 소급 보정 후 재고 추이가 표시됩니다.")
        self.chart_info.setObjectName("SectionHint")
        chart_frame_lay.addWidget(self.chart_info)
        chart_layout.addWidget(chart_frame, stretch=1)
        self.tabs.addTab(chart_page, "재고 추이 차트")

        # AI 분석 리포트
        ai_page = QWidget()
        ai_layout = QVBoxLayout(ai_page)
        ai_layout.setContentsMargins(8, 8, 8, 8)
        ai_layout.setSpacing(10)

        ai_top_card = QFrame()
        ai_top_card.setObjectName("Card")
        ai_top = QHBoxLayout(ai_top_card)
        ai_top.setContentsMargins(12, 10, 12, 10)
        self.ai_summary = QLabel("분석 대상: (데이터 없음)")
        self.ai_summary.setObjectName("SectionHint")
        self.analyze_btn = QPushButton("AI 분석 실행")
        self.analyze_btn.setObjectName("AccentButton")
        self.analyze_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.analyze_btn.clicked.connect(self.run_ai_analysis)
        ai_top.addWidget(self.ai_summary, stretch=1)
        ai_top.addWidget(self.analyze_btn)
        ai_layout.addWidget(ai_top_card)

        splitter = QSplitter(Qt.Orientation.Vertical)
        self.ai_prompt_view = QTextEdit()
        self.ai_prompt_view.setPlaceholderText("Gemini에 전달될 프롬프트(변화 품목만)가 표시됩니다.")
        self.ai_prompt_view.setReadOnly(True)
        self.ai_result_view = QTextEdit()
        self.ai_result_view.setPlaceholderText("AI 분석 리포트가 여기에 표시됩니다.")
        splitter.addWidget(self._wrap_card("전달 프롬프트 (변화 품목만)", self.ai_prompt_view))
        splitter.addWidget(self._wrap_card("AI 분석 리포트", self.ai_result_view))
        splitter.setSizes([280, 420])
        ai_layout.addWidget(splitter, stretch=1)
        self.tabs.addTab(ai_page, "AI 분석 리포트")

        # 하위 호환: 기존 테스트/호출이 path_label 을 참조할 수 있음
        self.path_label = self.status_file

    @staticmethod
    def _wrap_card(title: str, widget: QWidget) -> QWidget:
        box = QFrame()
        box.setObjectName("Card")
        lay = QVBoxLayout(box)
        lay.setContentsMargins(12, 10, 12, 10)
        lay.setSpacing(6)
        label = QLabel(title)
        label.setObjectName("FieldLabel")
        lay.addWidget(label)
        lay.addWidget(widget)
        return box

    def toggle_api_visibility(self) -> None:
        if self.api_key_edit.echoMode() == QLineEdit.EchoMode.Password:
            self.api_key_edit.setEchoMode(QLineEdit.EchoMode.Normal)
            self.api_toggle_btn.setText("숨김")
        else:
            self.api_key_edit.setEchoMode(QLineEdit.EchoMode.Password)
            self.api_toggle_btn.setText("표시")

    def current_api_key(self) -> str:
        return self.api_key_edit.text().strip()

    def refresh_api_status(self) -> None:
        if self.current_api_key():
            self.status_api.setText("API: 연결 준비됨")
            self.status_api.setStyleSheet(f"color: {COLORS['success']}; padding: 2px 10px;")
        else:
            self.status_api.setText("API: 미연결")
            self.status_api.setStyleSheet(f"color: {COLORS['warn']}; padding: 2px 10px;")

    def update_status_bar(self) -> None:
        if self._loaded_path:
            self.status_file.setText(f"파일: {os.path.basename(self._loaded_path)}")
            self.status_file.setToolTip(self._loaded_path)
        else:
            self.status_file.setText("파일: 없음")
            self.status_file.setToolTip("")
        self.status_correction.setText(f"소급 보정: {self._correction_count}건")
        self.refresh_api_status()

    @staticmethod
    def count_corrections(items: list[StockItem]) -> int:
        """동일일자 덮어쓰기로 소급 보정된 건수."""
        total = 0
        for it in items:
            if len(it.raw_points) > len(it.corrected_points):
                total += len(it.raw_points) - len(it.corrected_points)
        return total

    def open_excel(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "엑셀 파일 선택",
            os.path.join(os.path.dirname(__file__), "sample"),
            "Excel Files (*.xlsx *.xls)",
        )
        if path:
            self.load_file(path)

    def load_file(self, path: str) -> None:
        try:
            corrected_df, items = load_stock_excel(path)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "엑셀 로드 실패", str(exc))
            return

        self.corrected_df = corrected_df
        self.items = items
        self._item_by_manage_no = {it.manage_no: it for it in items if it.manage_no}
        self._loaded_path = path
        self._correction_count = self.count_corrections(items)
        self.drop_zone.set_file_name(path)
        self.update_status_bar()

        self.populate_table()
        self.populate_combo()
        self.refresh_ai_prompt_preview()
        if items:
            self.render_chart_for_item(items[0])

    def populate_table(self) -> None:
        df = self.corrected_df
        if df is None:
            self.table.clear()
            return

        sorting = self.table.isSortingEnabled()
        self.table.setSortingEnabled(False)
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels([str(c) for c in df.columns])

        for r, (_, row) in enumerate(df.iterrows()):
            manage_no = _cell_str(row.get("관리번호"))
            for c, col in enumerate(df.columns):
                value = row[col]
                if _is_empty(value):
                    text = ""
                elif isinstance(value, float) and float(value).is_integer():
                    text = str(int(value))
                else:
                    text = str(value)
                item = QTableWidgetItem(text)
                item.setData(Qt.ItemDataRole.UserRole, manage_no)
                self.table.setItem(r, c, item)
        self.table.resizeColumnsToContents()
        self.table.setSortingEnabled(sorting)

    def populate_combo(self) -> None:
        self._all_combo_labels = [(it.label, it.manage_no) for it in self.items]
        self.item_combo.blockSignals(True)
        self.item_combo.clear()
        for label, manage_no in self._all_combo_labels:
            self.item_combo.addItem(label, manage_no)
        if self._all_combo_labels:
            self.item_combo.setCurrentIndex(0)
            if self.item_combo.lineEdit() is not None:
                self.item_combo.lineEdit().setText(self._all_combo_labels[0][0])
        self.item_combo.blockSignals(False)

    def filter_combo_items(self, text: str) -> None:
        needle = text.strip().lower()
        current = self.item_combo.currentData()
        self.item_combo.blockSignals(True)
        self.item_combo.clear()
        for label, manage_no in self._all_combo_labels:
            if not needle or needle in label.lower() or needle in manage_no.lower():
                self.item_combo.addItem(label, manage_no)
        # 가능하면 이전 선택 유지
        if current:
            for i in range(self.item_combo.count()):
                if self.item_combo.itemData(i) == current:
                    self.item_combo.setCurrentIndex(i)
                    break
        self.item_combo.blockSignals(False)

    def on_table_double_clicked(self, row: int, _column: int) -> None:
        manage_no = ""
        cell = self.table.item(row, 0)
        if cell is not None:
            manage_no = cell.data(Qt.ItemDataRole.UserRole) or ""
        if not manage_no:
            headers = [
                self.table.horizontalHeaderItem(i).text()
                for i in range(self.table.columnCount())
            ]
            if "관리번호" in headers:
                idx = headers.index("관리번호")
                manage_item = self.table.item(row, idx)
                manage_no = manage_item.text().strip() if manage_item else ""

        if not manage_no or manage_no not in self._item_by_manage_no:
            QMessageBox.information(self, "알림", "해당 행의 관리번호를 확인할 수 없습니다.")
            return

        self.select_item_in_combo(manage_no)
        self.tabs.setCurrentIndex(1)  # 재고 추이 차트
        self.render_chart_for_item(self._item_by_manage_no[manage_no])

    def select_item_in_combo(self, manage_no: str) -> None:
        # 필터로 숨겨져 있을 수 있어 전체 목록 복원 후 선택
        self.item_combo.blockSignals(True)
        self.item_combo.clear()
        for label, code in self._all_combo_labels:
            self.item_combo.addItem(label, code)
        for i in range(self.item_combo.count()):
            if self.item_combo.itemData(i) == manage_no:
                self.item_combo.setCurrentIndex(i)
                if self.item_combo.lineEdit() is not None:
                    self.item_combo.lineEdit().setText(self.item_combo.itemText(i))
                break
        self.item_combo.blockSignals(False)

    def on_combo_changed(self, index: int) -> None:
        if index < 0:
            return
        manage_no = self.item_combo.itemData(index)
        item = self._item_by_manage_no.get(manage_no)
        if item:
            self.render_chart_for_item(item)

    def render_chart_for_item(self, item: StockItem) -> None:
        ax = self.canvas.ax
        ax.clear()
        ax.set_facecolor(COLORS["surface"])
        points = item.corrected_points
        if not points:
            ax.set_title(f"{item.label} — 변동 데이터 없음", color=COLORS["navy"], pad=12)
            ax.set_xlabel("변경일자", color=COLORS["muted"])
            ax.set_ylabel("재고량", color=COLORS["muted"])
            self.canvas.draw()
            self.chart_info.setText(f"{item.label}: 표시할 변경일자/재고량 쌍이 없습니다.")
            return

        xs = [format_date(p.change_date) for p in points]
        ys = [p.quantity for p in points]
        x_idx = list(range(len(ys)))
        ax.plot(
            x_idx,
            ys,
            marker="o",
            linewidth=2.4,
            markersize=7,
            color=COLORS["chart"],
            markerfacecolor=COLORS["accent"],
            markeredgecolor=COLORS["surface"],
            markeredgewidth=1.2,
        )
        ax.fill_between(x_idx, ys, color=COLORS["accent"], alpha=0.08)
        ax.set_xticks(x_idx)
        ax.set_xticklabels(xs, rotation=30, ha="right")
        ax.set_title(f"{item.label} 재고 추이 (소급 보정 후)", color=COLORS["navy"], pad=12)
        ax.set_xlabel("변경일자 (YYYY-MM-DD)", color=COLORS["muted"])
        ax.set_ylabel("재고량", color=COLORS["muted"])
        ax.grid(True, linestyle="--", alpha=0.28, color="#94A3B8")
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        self.canvas.draw()

        delta = item.qty_delta
        delta_text = f"{delta:+g}" if delta is not None else "-"
        self.chart_info.setText(
            f"{item.label}: 최초 {item.first_qty:g} → 최종 {item.last_qty:g} "
            f"(변화량 {delta_text}), 포인트 {len(points)}개"
        )

    def refresh_ai_prompt_preview(self) -> None:
        targets = items_for_ai_analysis(self.items)
        self.ai_summary.setText(
            f"분석 대상: 전체 {len(self.items)}건 중 수량 변화 품목 {len(targets)}건 "
            f"(변화 없음/데이터 없음 제외)"
        )
        if not targets:
            self.ai_prompt_view.setPlainText(
                "수량 변화가 있는 품목이 없어 AI 분석 대상이 없습니다."
            )
            return
        self.ai_prompt_view.setPlainText(build_ai_prompt(targets))

    def run_ai_analysis(self) -> None:
        targets = items_for_ai_analysis(self.items)
        if not targets:
            QMessageBox.information(
                self,
                "AI 분석",
                "수량 변화가 있는 품목이 없어 분석을 건너뜁니다.",
            )
            return

        api_key = self.current_api_key()
        if not api_key:
            QMessageBox.warning(
                self,
                "API 키 필요",
                "상단의 Gemini API Key 입력란에 키를 입력한 뒤 다시 실행하세요.",
            )
            return

        prompt = build_ai_prompt(targets)
        self.ai_prompt_view.setPlainText(prompt)
        self.ai_result_view.setPlainText("Gemini 분석 중...")
        self.status_api.setText("API: 분석 중...")
        self.status_api.setStyleSheet(f"color: {COLORS['blue']}; padding: 2px 10px;")

        if self._gemini_worker and self._gemini_worker.isRunning():
            QMessageBox.information(self, "AI 분석", "이미 분석이 진행 중입니다.")
            return

        self._gemini_worker = GeminiWorker(prompt, api_key, self)
        self._gemini_worker.finished_ok.connect(self._on_ai_ok)
        self._gemini_worker.failed.connect(self._on_ai_fail)
        self._gemini_worker.start()

    def _on_ai_ok(self, text: str) -> None:
        self.ai_result_view.setPlainText(text)
        self.status_api.setText("API: 연결됨")
        self.status_api.setStyleSheet(f"color: {COLORS['success']}; padding: 2px 10px;")

    def _on_ai_fail(self, err: str) -> None:
        self.ai_result_view.setPlainText(f"분석 실패: {err}")
        self.status_api.setText("API: 오류")
        self.status_api.setStyleSheet(f"color: {COLORS['danger']}; padding: 2px 10px;")


def main() -> int:
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(APP_STYLESHEET)
    font = QFont("Segoe UI", 10)
    app.setFont(font)

    window = StockApp()
    window.show()

    sample = os.path.join(os.path.dirname(__file__), "sample", "테스트파일1.xlsx")
    if os.path.exists(sample):
        window.load_file(sample)

    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
