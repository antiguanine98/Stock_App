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
from PySide6.QtGui import QAction
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSplitter,
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
        self.figure = Figure(figsize=(8, 4.5), tight_layout=True)
        super().__init__(self.figure)
        self.setParent(parent)
        self.ax = self.figure.add_subplot(111)


class StockApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("재고 소급 보정 · 추이 차트 · AI 분석")
        self.resize(1280, 820)

        self.items: list[StockItem] = []
        self.corrected_df: Optional[pd.DataFrame] = None
        self._item_by_manage_no: dict[str, StockItem] = {}
        self._gemini_worker: Optional[GeminiWorker] = None

        self._build_ui()
        self._build_menu()

    def _build_menu(self) -> None:
        open_action = QAction("엑셀 열기...", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.open_excel)
        file_menu = self.menuBar().addMenu("파일")
        file_menu.addAction(open_action)

    def _build_ui(self) -> None:
        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)

        toolbar = QHBoxLayout()
        self.path_label = QLabel("엑셀 파일을 열어 주세요. (예: sample/테스트파일1.xlsx)")
        self.path_label.setWordWrap(True)
        open_btn = QPushButton("엑셀 열기")
        open_btn.clicked.connect(self.open_excel)
        toolbar.addWidget(self.path_label, stretch=1)
        toolbar.addWidget(open_btn)
        layout.addLayout(toolbar)

        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)

        # --- 보정 데이터 표 ---
        table_page = QWidget()
        table_layout = QVBoxLayout(table_page)
        hint = QLabel("행의 셀을 더블클릭하면 해당 관리번호의 [재고 추이 차트]로 이동합니다.")
        hint.setStyleSheet("color: #555;")
        table_layout.addWidget(hint)
        self.table = QTableWidget()
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.cellDoubleClicked.connect(self.on_table_double_clicked)
        table_layout.addWidget(self.table)
        self.tabs.addTab(table_page, "보정 데이터 표")

        # --- 재고 추이 차트 ---
        chart_page = QWidget()
        chart_layout = QVBoxLayout(chart_page)
        selector_row = QHBoxLayout()
        selector_row.addWidget(QLabel("품목 선택 (관리번호(한글명)):"))
        self.item_combo = QComboBox()
        self.item_combo.currentIndexChanged.connect(self.on_combo_changed)
        selector_row.addWidget(self.item_combo, stretch=1)
        chart_layout.addLayout(selector_row)
        self.canvas = MplCanvas(self)
        chart_layout.addWidget(self.canvas)
        self.chart_info = QLabel("품목을 선택하면 소급 보정 후 재고 추이가 표시됩니다.")
        chart_layout.addWidget(self.chart_info)
        self.tabs.addTab(chart_page, "재고 추이 차트")

        # --- AI 분석 리포트 ---
        ai_page = QWidget()
        ai_layout = QVBoxLayout(ai_page)
        ai_top = QHBoxLayout()
        self.ai_summary = QLabel("분석 대상: (데이터 없음)")
        analyze_btn = QPushButton("AI 분석 실행")
        analyze_btn.clicked.connect(self.run_ai_analysis)
        ai_top.addWidget(self.ai_summary, stretch=1)
        ai_top.addWidget(analyze_btn)
        ai_layout.addLayout(ai_top)

        splitter = QSplitter(Qt.Orientation.Vertical)
        self.ai_prompt_view = QTextEdit()
        self.ai_prompt_view.setPlaceholderText("Gemini에 전달될 프롬프트(변화 품목만)가 표시됩니다.")
        self.ai_prompt_view.setReadOnly(True)
        self.ai_result_view = QTextEdit()
        self.ai_result_view.setPlaceholderText("AI 분석 리포트가 여기에 표시됩니다.")
        splitter.addWidget(self._wrap_labeled("전달 프롬프트 (변화 품목만)", self.ai_prompt_view))
        splitter.addWidget(self._wrap_labeled("AI 분석 리포트", self.ai_result_view))
        splitter.setSizes([280, 420])
        ai_layout.addWidget(splitter)
        self.tabs.addTab(ai_page, "AI 분석 리포트")

    @staticmethod
    def _wrap_labeled(title: str, widget: QWidget) -> QWidget:
        box = QWidget()
        lay = QVBoxLayout(box)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.addWidget(QLabel(title))
        lay.addWidget(widget)
        return box

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
        self.path_label.setText(f"로드됨: {path}  |  품목 {len(items)}건")

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

    def populate_combo(self) -> None:
        self.item_combo.blockSignals(True)
        self.item_combo.clear()
        for it in self.items:
            self.item_combo.addItem(it.label, it.manage_no)
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
        for i in range(self.item_combo.count()):
            if self.item_combo.itemData(i) == manage_no:
                self.item_combo.blockSignals(True)
                self.item_combo.setCurrentIndex(i)
                self.item_combo.blockSignals(False)
                return

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
        points = item.corrected_points
        if not points:
            ax.set_title(f"{item.label} — 변동 데이터 없음")
            ax.set_xlabel("변경일자")
            ax.set_ylabel("재고량")
            self.canvas.draw()
            self.chart_info.setText(f"{item.label}: 표시할 변경일자/재고량 쌍이 없습니다.")
            return

        xs = [format_date(p.change_date) for p in points]
        ys = [p.quantity for p in points]
        ax.plot(xs, ys, marker="o", linewidth=2, color="#1f6f8b")
        ax.set_title(f"{item.label} 재고 추이 (소급 보정 후)")
        ax.set_xlabel("변경일자 (YYYY-MM-DD)")
        ax.set_ylabel("재고량")
        ax.grid(True, linestyle="--", alpha=0.35)
        for label in ax.get_xticklabels():
            label.set_rotation(30)
            label.set_ha("right")
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

        api_key = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")
        if not api_key:
            QMessageBox.warning(
                self,
                "API 키 필요",
                "환경변수 GEMINI_API_KEY(또는 GOOGLE_API_KEY)를 설정한 뒤 다시 실행하세요.",
            )
            return

        prompt = build_ai_prompt(targets)
        self.ai_prompt_view.setPlainText(prompt)
        self.ai_result_view.setPlainText("Gemini 분석 중...")

        if self._gemini_worker and self._gemini_worker.isRunning():
            QMessageBox.information(self, "AI 분석", "이미 분석이 진행 중입니다.")
            return

        self._gemini_worker = GeminiWorker(prompt, api_key, self)
        self._gemini_worker.finished_ok.connect(self.ai_result_view.setPlainText)
        self._gemini_worker.failed.connect(
            lambda err: self.ai_result_view.setPlainText(f"분석 실패: {err}")
        )
        self._gemini_worker.start()


def main() -> int:
    app = QApplication(sys.argv)
    window = StockApp()
    window.show()

    sample = os.path.join(os.path.dirname(__file__), "sample", "테스트파일1.xlsx")
    if os.path.exists(sample):
        window.load_file(sample)

    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
