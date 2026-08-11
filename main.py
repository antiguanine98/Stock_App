"""
생약표준품 재고 분석 및 소급 보정 시스템 (PyQt6) v1.11
"""

from __future__ import annotations

import json
import sys
import traceback
from pathlib import Path
from typing import Any

import matplotlib
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from PyQt6.QtCore import QEvent, QObject, Qt, QThread, QTimer, QUrl, pyqtSignal
from PyQt6.QtGui import (
    QBrush,
    QColor,
    QDragEnterEvent,
    QDropEvent,
    QFont,
    QPainter,
    QPainterPath,
    QPen,
    QPixmap,
)
from PyQt6.QtWidgets import (
    QApplication,
    QComboBox,
    QCompleter,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSplashScreen,
    QStatusBar,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from stock_logic import (
    StockItem,
    build_ai_prompt,
    build_scatter3d_records,
    process_excel,
)

matplotlib.use("QtAgg")
matplotlib.rcParams["font.family"] = "Malgun Gothic"
matplotlib.rcParams["axes.unicode_minus"] = False

try:
    import mplcursors
except ImportError:  # pragma: no cover
    mplcursors = None

try:
    from PyQt6.QtWebEngineWidgets import QWebEngineView
except ImportError:  # pragma: no cover
    QWebEngineView = None  # type: ignore[misc, assignment]

CONFIG_PATH = Path(__file__).parent / "config.json"
VIEWER_HTML_PATH = Path(__file__).parent / "viewer.html"
APP_VERSION = "v1.11"
AUTHOR_CREDIT = "made by 2026MFDSyouthinternKYHLCY"

GEMINI_MODEL_PREFERENCES = [
    "gemini-3.1-flash-lite",
    "gemini-flash-latest",
    "gemini-3.5-flash-lite",
    "gemini-3.6-flash",
    "gemini-3.5-flash",
    "gemini-2.5-flash-lite",
    "gemini-2.5-flash",
]
_resolved_gemini_model: str | None = None
_MAX_MODEL_PROBES = 4

APP_STYLE = """
QMainWindow, QWidget {
    background-color: #f4f6f9;
    color: #1a2332;
    font-family: "Malgun Gothic", "Segoe UI", sans-serif;
    font-size: 13px;
}
QLabel#titleLabel {
    color: #0b1f3a;
    font-size: 18px;
    font-weight: 700;
}
QLabel#sectionLabel {
    color: #1e3a5f;
    font-size: 12px;
    font-weight: 700;
}
QFrame#card {
    background-color: #ffffff;
    border: 1px solid #d8e0ea;
    border-radius: 12px;
}
QFrame#dropZone {
    border: 2px dashed #3d5a80;
    border-radius: 12px;
    background-color: #eef3f9;
}
QFrame#dropZone:hover {
    border-color: #1b3a6b;
    background-color: #e4edf8;
}
QLineEdit, QTextEdit {
    background-color: #ffffff;
    border: 1px solid #c5d0de;
    border-radius: 8px;
    padding: 8px 10px;
    selection-background-color: #1e3a5f;
}
QLineEdit:focus, QTextEdit:focus {
    border: 1px solid #1e3a5f;
}
QComboBox {
    background-color: #ffffff;
    border: 1px solid #c5d0de;
    border-radius: 8px;
    padding: 4px 8px 4px 10px;
    min-height: 32px;
    selection-background-color: #1e3a5f;
}
QComboBox:focus, QComboBox:on {
    border: 1px solid #1e3a5f;
}
QComboBox::drop-down {
    subcontrol-origin: padding;
    subcontrol-position: center right;
    width: 34px;
    border: none;
    border-left: 1px solid #d0dae8;
    background-color: #eef3f9;
    border-top-right-radius: 8px;
    border-bottom-right-radius: 8px;
}
QComboBox::drop-down:hover {
    background-color: #dce6f2;
}
QComboBox::down-arrow {
    width: 0;
    height: 0;
    border-left: 5px solid transparent;
    border-right: 5px solid transparent;
    border-top: 6px solid #1e3a5f;
    margin-right: 2px;
}
QComboBox QAbstractItemView {
    background-color: #ffffff;
    border: 1px solid #c5d0de;
    selection-background-color: #d6e4f7;
    selection-color: #0b1f3a;
    outline: 0;
    padding: 4px;
}
QPushButton {
    background-color: #1e3a5f;
    color: #ffffff;
    border: none;
    border-radius: 8px;
    padding: 9px 16px;
    font-weight: 600;
}
QPushButton:hover { background-color: #274b7a; }
QPushButton:pressed { background-color: #152a45; }
QPushButton:disabled { background-color: #94a3b8; }
QPushButton#secondaryBtn {
    background-color: #ffffff;
    color: #1e3a5f;
    border: 1px solid #1e3a5f;
}
QPushButton#secondaryBtn:hover { background-color: #eef3f9; }
QTabWidget::pane {
    border: 1px solid #d8e0ea;
    border-radius: 12px;
    background: #ffffff;
    top: -1px;
}
QTabBar::tab {
    background: #e8eef6;
    color: #3a4d66;
    border: 1px solid #d0dae8;
    border-bottom: none;
    border-top-left-radius: 8px;
    border-top-right-radius: 8px;
    padding: 9px 16px;
    margin-right: 4px;
    min-width: 96px;
}
QTabBar::tab:selected {
    background: #ffffff;
    color: #0b1f3a;
    font-weight: 700;
}
QTabBar::tab:hover:!selected { background: #dce6f2; }
QTableWidget {
    background-color: #ffffff;
    alternate-background-color: #f7f9fc;
    gridline-color: #e3eaf3;
    border: none;
    selection-background-color: #d6e4f7;
    selection-color: #0b1f3a;
}
QHeaderView::section {
    background-color: #1e3a5f;
    color: #ffffff;
    padding: 8px;
    border: none;
    font-weight: 600;
}
QStatusBar {
    background-color: #ffffff;
    border-top: 1px solid #d8e0ea;
}
QStatusBar QLabel { color: #334155; padding: 2px 10px; }
QLabel#creditLabel {
    color: #94a3b8;
    font-size: 11px;
    padding: 2px 12px;
}
QStatusBar QFrame#statusDivider {
    background-color: #d0dae8;
    max-width: 1px;
    margin: 4px 2px;
}
"""


# ---------------------------------------------------------------------------
# Gemini helpers
# ---------------------------------------------------------------------------


def log_gemini(level: str, message: str, exc: BaseException | None = None) -> None:
    print(f"[Gemini {level}] {message}", file=sys.stderr, flush=True)
    if exc is not None:
        print(traceback.format_exc(), file=sys.stderr, flush=True)


def format_gemini_error(exc: BaseException) -> str:
    try:
        from google.genai import errors as genai_errors

        if isinstance(exc, genai_errors.APIError):
            code = getattr(exc, "code", "?")
            status = getattr(exc, "status", "") or "UNKNOWN"
            message = getattr(exc, "message", None) or str(exc)
            return f"HTTP {code} {status}: {message}"
    except ImportError:
        pass
    return f"{type(exc).__name__}: {exc}"


def create_gemini_client(api_key: str):
    from google import genai

    return genai.Client(api_key=api_key)


def _api_error_code(exc: BaseException) -> int | None:
    try:
        from google.genai import errors as genai_errors

        if isinstance(exc, genai_errors.APIError):
            code = getattr(exc, "code", None)
            return int(code) if code is not None else None
    except (ImportError, TypeError, ValueError):
        pass
    return None


def _is_auth_error(exc: BaseException) -> bool:
    code = _api_error_code(exc)
    if code in (401, 403):
        return True
    text = str(exc).lower()
    return any(k in text for k in ("api key", "api_key", "permission denied", "unauthenticated"))


def _is_model_unavailable_error(exc: BaseException) -> bool:
    text = str(exc).lower()
    markers = (
        "no longer available", "not found", "not_found", "unsupported",
        "resource_exhausted", "rate limit", "quota", "deadline", "timed out", "timeout",
    )
    if any(m in text for m in markers):
        return True
    code = _api_error_code(exc)
    return code in (404, 400, 429, 503, 504)


def _candidate_models(prioritize: str | None = None) -> list[str]:
    ordered: list[str] = []
    if prioritize:
        ordered.append(prioritize)
    if _resolved_gemini_model and _resolved_gemini_model not in ordered:
        ordered.append(_resolved_gemini_model)
    for preferred in GEMINI_MODEL_PREFERENCES:
        if preferred not in ordered:
            ordered.append(preferred)
    return ordered[:_MAX_MODEL_PROBES]


def extract_response_text(response: Any) -> str | None:
    text = getattr(response, "text", None)
    if text and str(text).strip():
        return str(text).strip()
    return None


def _probe_model(client, model: str) -> bool:
    from google.genai import types

    try:
        response = client.models.generate_content(
            model=model,
            contents="Reply with exactly one word: OK",
            config=types.GenerateContentConfig(max_output_tokens=16, temperature=0),
        )
        return extract_response_text(response) is not None
    except Exception as exc:
        if _is_auth_error(exc):
            raise
        log_gemini("WARN", f"모델 사용 불가, 건너뜀: {model} ({format_gemini_error(exc)})")
        return False


def resolve_gemini_model(client, force_refresh: bool = False) -> str:
    global _resolved_gemini_model
    if _resolved_gemini_model and not force_refresh:
        return _resolved_gemini_model

    errors: list[str] = []
    for model in _candidate_models(prioritize=_resolved_gemini_model):
        log_gemini("INFO", f"모델 후보 검증: {model}")
        try:
            if _probe_model(client, model):
                _resolved_gemini_model = model
                log_gemini("INFO", f"사용 모델 선택: {model}")
                return model
            errors.append(f"{model}: 응답 없음 또는 사용 불가")
        except Exception as exc:
            if _is_auth_error(exc):
                raise
            detail = format_gemini_error(exc)
            errors.append(f"{model}: {detail}")
            log_gemini("ERROR", f"모델 검증 실패: {model} — {detail}", exc)

    raise RuntimeError("사용 가능한 Gemini Flash 모델을 찾지 못했습니다.\n" + "\n".join(errors[:8]))


def get_active_gemini_model() -> str:
    return _resolved_gemini_model or GEMINI_MODEL_PREFERENCES[0]


def describe_empty_response(response: Any) -> str:
    parts = ["모델 응답 본문이 비어 있습니다."]
    prompt_feedback = getattr(response, "prompt_feedback", None)
    if prompt_feedback and getattr(prompt_feedback, "block_reason", None):
        parts.append(f"프롬프트 차단 사유: {prompt_feedback.block_reason}")
    candidates = getattr(response, "candidates", None) or []
    if not candidates:
        parts.append("candidates가 없습니다.")
    else:
        finish_reason = getattr(candidates[0], "finish_reason", None)
        if finish_reason:
            parts.append(f"finish_reason: {finish_reason}")
    return " ".join(parts)


def test_gemini_connection(api_key: str) -> tuple[bool, str]:
    global _resolved_gemini_model
    client = create_gemini_client(api_key)
    log_gemini("INFO", "연결 테스트 시작")
    if _resolved_gemini_model:
        if _probe_model(client, _resolved_gemini_model):
            return True, f"연결됨 ({_resolved_gemini_model})"
        _resolved_gemini_model = None
    model = resolve_gemini_model(client, force_refresh=True)
    return True, f"연결됨 ({model})"


def generate_gemini_report(api_key: str, prompt: str) -> str:
    global _resolved_gemini_model
    client = create_gemini_client(api_key)
    model = resolve_gemini_model(client)
    log_gemini("INFO", f"AI 분석 요청 시작 (model={model})")
    try:
        response = client.models.generate_content(model=model, contents=prompt)
    except Exception as exc:
        if _is_model_unavailable_error(exc):
            _resolved_gemini_model = None
            model = resolve_gemini_model(client, force_refresh=True)
            response = client.models.generate_content(model=model, contents=prompt)
        else:
            raise
    text = extract_response_text(response)
    if text:
        return text
    raise RuntimeError(describe_empty_response(response))


def load_config() -> dict[str, Any]:
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            pass
    return {}


def save_config(data: dict[str, Any]) -> None:
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# Workers
# ---------------------------------------------------------------------------


class ExcelWorker(QThread):
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, file_path: str):
        super().__init__()
        self.file_path = file_path

    def run(self) -> None:
        try:
            self.finished.emit(process_excel(self.file_path))
        except Exception as e:
            self.error.emit(str(e))


class GeminiWorker(QThread):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, api_key: str, prompt: str):
        super().__init__()
        self.api_key = api_key
        self.prompt = prompt

    def run(self) -> None:
        try:
            self.finished.emit(generate_gemini_report(self.api_key, self.prompt))
        except Exception as e:
            detail = format_gemini_error(e)
            log_gemini("ERROR", f"AI 분석 실패: {detail}", e)
            self.error.emit(detail)


class GeminiTestWorker(QThread):
    finished = pyqtSignal(bool, str)

    def __init__(self, api_key: str):
        super().__init__()
        self.api_key = api_key

    def run(self) -> None:
        try:
            ok, message = test_gemini_connection(self.api_key)
            self.finished.emit(ok, message)
        except Exception as e:
            detail = format_gemini_error(e)
            log_gemini("ERROR", f"연결 테스트 실패: {detail}", e)
            self.finished.emit(False, detail)


# ---------------------------------------------------------------------------
# UI widgets
# ---------------------------------------------------------------------------


class ComboPopupFilter(QObject):
    def __init__(self, combo: QComboBox):
        super().__init__(combo)
        self._combo = combo

    def eventFilter(self, obj: QObject, event: QEvent) -> bool:
        if event.type() == QEvent.Type.MouseButtonPress:
            self._combo.showPopup()
        return False


class DropZone(QFrame):
    file_dropped = pyqtSignal(str)

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setMinimumHeight(96)
        self.setObjectName("dropZone")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 14)
        title = QLabel("엑셀 파일을 여기에 드래그 앤 드롭")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet(
            "color: #1e3a5f; font-size: 14px; font-weight: 700; border: none; background: transparent;"
        )
        subtitle = QLabel(".xlsx / .xls 지원 · 또는 오른쪽 [파일 선택] 버튼 사용")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setStyleSheet(
            "color: #5b6b7c; font-size: 12px; border: none; background: transparent;"
        )
        layout.addStretch()
        layout.addWidget(title)
        layout.addWidget(subtitle)
        layout.addStretch()

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            url = event.mimeData().urls()[0]
            if url.toLocalFile().lower().endswith((".xlsx", ".xls")):
                event.acceptProposedAction()
                return
        event.ignore()

    def dropEvent(self, event: QDropEvent) -> None:
        path = event.mimeData().urls()[0].toLocalFile()
        if path.lower().endswith((".xlsx", ".xls")):
            self.file_dropped.emit(path)
            event.acceptProposedAction()
        else:
            event.ignore()


class InventoryChart(FigureCanvas):
    def __init__(self, parent: QWidget | None = None):
        self.figure = Figure(figsize=(10, 5), dpi=100)
        self.figure.set_facecolor("#ffffff")
        super().__init__(self.figure)
        self.setParent(parent)
        self._annot = None
        self._hover_cid = None
        self._cursor = None
        self._points: list[tuple[float, float, str]] = []
        self.mpl_connect("motion_notify_event", self._on_hover)
        self._show_placeholder("품목 또는 표준품구분을 선택하면 재고 추이 차트가 표시됩니다.")

    def _clear_helpers(self) -> None:
        self._annot = None
        self._cursor = None
        self._points = []

    def _show_placeholder(self, message: str) -> None:
        self._clear_helpers()
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        ax.set_facecolor("#fafbfc")
        ax.text(0.5, 0.5, message, ha="center", va="center", color="#64748b", fontsize=12)
        ax.set_xticks([])
        ax.set_yticks([])
        for spine in ax.spines.values():
            spine.set_visible(False)
        self.figure.tight_layout()
        self.draw()

    def _attach_hover(self, ax, line, labels: list[str]) -> None:
        xdata = list(line.get_xdata())
        ydata = list(line.get_ydata())
        self._points = [
            (float(x), float(y), labels[i] if i < len(labels) else "")
            for i, (x, y) in enumerate(zip(xdata, ydata))
            if y == y  # not NaN
        ]
        self._annot = ax.annotate(
            "",
            xy=(0, 0),
            xytext=(12, 12),
            textcoords="offset points",
            bbox=dict(boxstyle="round,pad=0.35", fc="#0b1f3a", ec="none", alpha=0.92),
            color="white",
            fontsize=9,
        )
        self._annot.set_visible(False)

        if mplcursors is not None:
            self._cursor = mplcursors.cursor(line, hover=True)

            @self._cursor.connect("add")
            def _on_add(sel):  # type: ignore[no-untyped-def]
                idx = int(sel.index) if sel.index is not None else 0
                if 0 <= idx < len(labels):
                    sel.annotation.set_text(labels[idx])
                sel.annotation.get_bbox_patch().set(fc="#0b1f3a", alpha=0.92)
                sel.annotation.set_color("white")

    def _on_hover(self, event) -> None:
        if mplcursors is not None or self._annot is None or not self._points:
            return
        if event.inaxes is None or event.xdata is None or event.ydata is None:
            if self._annot.get_visible():
                self._annot.set_visible(False)
                self.draw_idle()
            return
        # 가장 가까운 포인트
        nearest = min(
            self._points,
            key=lambda p: (p[0] - event.xdata) ** 2 + (p[1] - event.ydata) ** 2,
        )
        dist = ((nearest[0] - event.xdata) ** 2 + (nearest[1] - event.ydata) ** 2) ** 0.5
        if dist < 0.35:
            self._annot.xy = (nearest[0], nearest[1])
            self._annot.set_text(nearest[2])
            self._annot.set_visible(True)
            self.draw_idle()
        elif self._annot.get_visible():
            self._annot.set_visible(False)
            self.draw_idle()

    def plot_item(
        self,
        label: str,
        dates: list[str],
        corrected: list[float | None],
        original: list[float | None] | None = None,
        original_dates: list[str] | None = None,
    ) -> None:
        if not dates:
            self._show_placeholder(f"{label}\n표시할 변동기록이 없습니다.")
            return

        self._clear_helpers()
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        ax.set_facecolor("#fafbfc")
        x = list(range(len(dates)))
        y_corr = [v if v is not None else float("nan") for v in corrected]

        if original is not None and original_dates is not None and original_dates == dates:
            y_orig = [v if v is not None else float("nan") for v in original]
            ax.plot(x, y_orig, "o--", label="연도말 원본", color="#c0392b", linewidth=1.5, markersize=5, alpha=0.7)

        (line,) = ax.plot(
            x, y_corr, "s-", label="연도말 소급 보정",
            color="#1e3a5f", linewidth=2.2, markersize=7,
        )
        labels = [
            f"{label}\n일자: {d}\n재고량: {y:g}" if y == y else f"{label}\n일자: {d}"
            for d, y in zip(dates, y_corr)
        ]
        self._attach_hover(ax, line, labels)

        ax.set_xticks(x)
        ax.set_xticklabels(dates, rotation=35, ha="right", fontsize=8)
        ax.set_xlabel("변경일자 (연도별 최종)")
        ax.set_ylabel("재고량")
        ax.set_title(f"재고 추이 — {label}", fontsize=12, fontweight="bold", color="#0b1f3a", pad=12)
        ax.legend(loc="best", frameon=False)
        ax.grid(True, axis="y", alpha=0.28, linestyle="--")
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        self.figure.tight_layout()
        self.draw()

    def plot_category(self, category: str, items: list[dict[str, Any]]) -> None:
        valid = [it for it in items if it.get("dates")]
        if not valid:
            self._show_placeholder(f"[{category}] 표시할 품목이 없습니다.")
            return

        self._clear_helpers()
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        ax.set_facecolor("#fafbfc")
        palette = ["#1e3a5f", "#2563eb", "#0f766e", "#b45309", "#be123c", "#7c3aed", "#0369a1"]
        hover_labels: list[str] = []
        lines = []

        # 비교를 위해 최대 12개만 표시
        for i, item in enumerate(valid[:12]):
            dates = item["dates"]
            y = [v if v is not None else float("nan") for v in item["corrected"]]
            x = list(range(len(dates)))
            (line,) = ax.plot(
                x, y, "o-",
                label=item["label"][:28],
                color=palette[i % len(palette)],
                linewidth=1.8,
                markersize=5,
            )
            lines.append(line)
            for d, q in zip(dates, y):
                hover_labels.append(
                    f"{item['label']}\n일자: {d}\n재고량: {q:g}" if q == q else f"{item['label']}\n일자: {d}"
                )

        if lines:
            # 대표 라인에 hover (mplcursors는 복수 아티스트 지원)
            if mplcursors is not None:
                self._cursor = mplcursors.cursor(lines, hover=True)

                @self._cursor.connect("add")
                def _on_add(sel):  # type: ignore[no-untyped-def]
                    artist = sel.artist
                    idx = lines.index(artist) if artist in lines else 0
                    item = valid[idx]
                    di = int(sel.index) if sel.index is not None else 0
                    if 0 <= di < len(item["dates"]):
                        q = item["corrected"][di]
                        qtxt = f"{q:g}" if q is not None else "-"
                        sel.annotation.set_text(
                            f"{item['label']}\n일자: {item['dates'][di]}\n재고량: {qtxt}"
                        )
                    sel.annotation.get_bbox_patch().set(fc="#0b1f3a", alpha=0.92)
                    sel.annotation.set_color("white")
            else:
                self._attach_hover(ax, lines[0], hover_labels)

        ax.set_xlabel("시점 인덱스 (품목별 연도말 순서)")
        ax.set_ylabel("재고량")
        ax.set_title(f"카테고리 비교 — {category}", fontsize=12, fontweight="bold", color="#0b1f3a")
        ax.legend(loc="best", fontsize=8, frameon=False)
        ax.grid(True, axis="y", alpha=0.28, linestyle="--")
        self.figure.tight_layout()
        self.draw()


class Scatter3DView(QWidget):
    """viewer.html 기반 3D 산점도 (QWebEngineView)."""

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setObjectName("scatter3dHost")
        self.setStyleSheet(
            "#scatter3dHost { background-color: #0f1612; border: 1px solid #24332b; border-radius: 8px; }"
        )
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self._fallback = QLabel(
            "엑셀을 로드하면 총 변동량 · 잔존 예상 소진기간 · 연평균 분양량 3D 산점도가 표시됩니다."
        )
        self._fallback.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._fallback.setWordWrap(True)
        self._fallback.setStyleSheet("color: #B9CCC0; font-size: 13px; padding: 24px;")
        layout.addWidget(self._fallback)

        self._web: Any = None
        if QWebEngineView is not None:
            try:
                web = QWebEngineView(self)
                web.setContextMenuPolicy(Qt.ContextMenuPolicy.NoContextMenu)
                web.page().setBackgroundColor(QColor(15, 22, 18))
                layout.addWidget(web, stretch=1)
                self._web = web
                self._fallback.hide()
                self.show_message(
                    "엑셀을 로드하면 총 변동량 · 잔존 예상 소진기간 · 연평균 분양량 3D 산점도가 표시됩니다."
                )
            except Exception:
                self._web = None
                self._fallback.setText(
                    "3D 웹뷰를 초기화하지 못했습니다. PyQt6-WebEngine 설치를 확인해 주세요."
                )
                self._fallback.show()
        else:
            self._fallback.setText(
                "3D 산점도를 표시하려면 PyQt6-WebEngine이 필요합니다.\n"
                "pip install PyQt6-WebEngine 후 다시 실행해 주세요."
            )

    def _load_template(self) -> str:
        if not VIEWER_HTML_PATH.exists():
            raise FileNotFoundError(f"viewer.html을 찾을 수 없습니다: {VIEWER_HTML_PATH}")
        return VIEWER_HTML_PATH.read_text(encoding="utf-8")

    def _set_html(self, html: str) -> None:
        if self._web is None:
            return
        base = QUrl.fromLocalFile(str(VIEWER_HTML_PATH.resolve()))
        self._web.setHtml(html, base)
        self._web.show()
        self._fallback.hide()

    def show_message(self, message: str) -> None:
        """탭 안에 안내 메시지용 최소 HTML을 표시 (크래시 없이)."""
        if self._web is None:
            self._fallback.setText(message)
            self._fallback.show()
            return
        safe = (
            message.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace("\n", "<br>")
        )
        html = f"""<!DOCTYPE html><html><head><meta charset="UTF-8">
<style>
html,body{{margin:0;height:100%;background:transparent;color:#B9CCC0;
font-family:-apple-system,'Malgun Gothic',sans-serif;}}
.wrap{{display:flex;align-items:center;justify-content:center;height:100%;
padding:28px;text-align:center;line-height:1.7;font-size:14px;}}
</style></head><body><div class="wrap">{safe}</div></body></html>"""
        self._set_html(html)

    def plot_records(self, records: list[dict[str, Any]], source_file: str = "") -> None:
        if self._web is None:
            self._fallback.setText(
                f"3D 레코드 {len(records)}건 — PyQt6-WebEngine이 없어 렌더링할 수 없습니다."
            )
            self._fallback.show()
            return
        try:
            template = self._load_template()
        except Exception as exc:
            self.show_message(f"viewer.html을 읽지 못했습니다.\n{exc}")
            return

        html = template.replace(
            "/*__RECORDS__*/[]",
            json.dumps(records, ensure_ascii=False),
        ).replace(
            '/*__SOURCE_FILE__*/""',
            json.dumps(source_file or "", ensure_ascii=False),
        )
        self._set_html(html)


# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"생약표준품 재고 분석 시스템 {APP_VERSION}")
        self.setMinimumSize(1180, 780)
        self.resize(1360, 880)

        self.inventory_data: dict[str, Any] | None = None
        self.excel_worker: ExcelWorker | None = None
        self.gemini_worker: GeminiWorker | None = None
        self.test_worker: GeminiTestWorker | None = None
        self._item_by_label: dict[str, dict[str, Any]] = {}
        self._updating_combo = False

        self._build_ui()
        self._load_api_key()

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setSpacing(12)
        root.setContentsMargins(18, 16, 18, 8)

        title = QLabel(f"생약표준품 재고 분석 시스템 {APP_VERSION}")
        title.setObjectName("titleLabel")
        root.addWidget(title)

        top_card = QFrame()
        top_card.setObjectName("card")
        top_layout = QVBoxLayout(top_card)
        top_layout.setContentsMargins(14, 12, 14, 12)
        top_layout.setSpacing(10)

        file_label = QLabel("엑셀 업로드")
        file_label.setObjectName("sectionLabel")
        top_layout.addWidget(file_label)

        file_row = QHBoxLayout()
        self.dropzone = DropZone()
        self.dropzone.file_dropped.connect(self._load_excel)
        file_row.addWidget(self.dropzone, stretch=1)
        btn_select = QPushButton("파일 선택")
        btn_select.setObjectName("secondaryBtn")
        btn_select.setMinimumWidth(110)
        btn_select.setMinimumHeight(42)
        btn_select.clicked.connect(self._browse_file)
        file_row.addWidget(btn_select, alignment=Qt.AlignmentFlag.AlignVCenter)
        top_layout.addLayout(file_row)

        api_label = QLabel("Gemini API Key")
        api_label.setObjectName("sectionLabel")
        top_layout.addWidget(api_label)

        api_row = QHBoxLayout()
        self.api_key_input = QLineEdit()
        self.api_key_input.setPlaceholderText("Google Gemini API Key 입력")
        self.api_key_input.setEchoMode(QLineEdit.EchoMode.Password)
        api_row.addWidget(self.api_key_input, stretch=1)
        btn_save_key = QPushButton("저장")
        btn_save_key.clicked.connect(self._save_api_key)
        api_row.addWidget(btn_save_key)
        self.btn_test = QPushButton("연결 테스트")
        self.btn_test.setObjectName("secondaryBtn")
        self.btn_test.clicked.connect(self._test_api_connection)
        api_row.addWidget(self.btn_test)
        top_layout.addLayout(api_row)
        root.addWidget(top_card)

        self.tabs = QTabWidget()

        # Tab 1
        table_wrap = QWidget()
        table_layout = QVBoxLayout(table_wrap)
        table_layout.setContentsMargins(8, 8, 8, 8)
        hint = QLabel("행을 더블클릭하면 해당 품목의 재고 추이 차트로 이동합니다. (등록일자 제외 · 잔고→재고 표기)")
        hint.setStyleSheet("color: #64748b; font-size: 12px;")
        table_layout.addWidget(hint)
        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(False)
        self.table.cellDoubleClicked.connect(self._on_table_double_clicked)
        table_layout.addWidget(self.table)
        self.tabs.addTab(table_wrap, "보정 데이터 표")

        # Tab 2 chart
        chart_widget = QWidget()
        chart_layout = QVBoxLayout(chart_widget)
        chart_layout.setContentsMargins(10, 10, 10, 10)
        chart_layout.setSpacing(8)

        filter_row = QHBoxLayout()
        cat_label = QLabel("표준품구분")
        cat_label.setObjectName("sectionLabel")
        filter_row.addWidget(cat_label)
        self.category_combo = QComboBox()
        self.category_combo.addItem("전체")
        self.category_combo.setMinimumWidth(160)
        self.category_combo.currentIndexChanged.connect(self._on_category_changed)
        filter_row.addWidget(self.category_combo)

        item_label = QLabel("품목 선택")
        item_label.setObjectName("sectionLabel")
        filter_row.addWidget(item_label)
        self.item_combo = QComboBox()
        self.item_combo.setEditable(True)
        self.item_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.item_combo.setMinimumWidth(300)
        self.item_combo.setMaxVisibleItems(20)
        line_edit = self.item_combo.lineEdit()
        if line_edit is not None:
            line_edit.setPlaceholderText("관리번호(한글명) 검색/선택")
            line_edit.setClearButtonEnabled(True)
            self._combo_popup_filter = ComboPopupFilter(self.item_combo)
            line_edit.installEventFilter(self._combo_popup_filter)
        completer = QCompleter(self.item_combo.model(), self.item_combo)
        completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        completer.setFilterMode(Qt.MatchFlag.MatchContains)
        completer.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
        self.item_combo.setCompleter(completer)
        self.item_combo.currentIndexChanged.connect(self._on_item_combo_changed)
        filter_row.addWidget(self.item_combo, stretch=1)
        btn_open_list = QPushButton("목록")
        btn_open_list.setObjectName("secondaryBtn")
        btn_open_list.setFixedWidth(72)
        btn_open_list.clicked.connect(self.item_combo.showPopup)
        filter_row.addWidget(btn_open_list)
        chart_layout.addLayout(filter_row)

        hover_hint = QLabel("그래프 포인트에 마우스를 올리면 일자·재고량이 표시됩니다.")
        hover_hint.setStyleSheet("color: #64748b; font-size: 12px;")
        chart_layout.addWidget(hover_hint)

        self.chart = InventoryChart()
        chart_layout.addWidget(self.chart, stretch=1)
        self.tabs.addTab(chart_widget, "재고 추이 차트")

        # Tab 3 AI
        report_wrap = QWidget()
        report_layout = QVBoxLayout(report_wrap)
        report_layout.setContentsMargins(8, 8, 8, 8)
        self.report_edit = QTextEdit()
        self.report_edit.setReadOnly(True)
        self.report_edit.setPlaceholderText(
            "엑셀 업로드 후 생약표준품 분양·소진 예측 AI 리포트가 생성됩니다."
        )
        self.report_edit.setFont(QFont("Malgun Gothic", 10))
        report_layout.addWidget(self.report_edit)
        self.tabs.addTab(report_wrap, "AI 분석 리포트")

        # Tab 4 3D scatter (viewer.html)
        viz_wrap = QWidget()
        viz_layout = QVBoxLayout(viz_wrap)
        viz_layout.setContentsMargins(10, 10, 10, 10)
        viz_row = QHBoxLayout()
        viz_label = QLabel("표준품구분 필터")
        viz_label.setObjectName("sectionLabel")
        viz_row.addWidget(viz_label)
        self.viz_category_combo = QComboBox()
        self.viz_category_combo.addItem("전체")
        self.viz_category_combo.currentIndexChanged.connect(self._refresh_3d)
        viz_row.addWidget(self.viz_category_combo)
        viz_row.addStretch(1)
        viz_layout.addLayout(viz_row)
        viz_hint = QLabel(
            "드래그로 회전 · 휠/핀치로 확대축소 · 호버/탭으로 품목 정보 "
            "(X: 총 변동량 · Y: 잔존 예상 소진기간 · Z: 연평균 분양량)"
        )
        viz_hint.setStyleSheet("color: #64748b; font-size: 12px;")
        viz_layout.addWidget(viz_hint)
        self.chart3d = Scatter3DView()
        viz_layout.addWidget(self.chart3d, stretch=1)
        self.tabs.addTab(viz_wrap, "3D/통합 시각화")

        root.addWidget(self.tabs, stretch=1)

        status = QStatusBar()
        self.setStatusBar(status)
        self.status_excel = QLabel("파일: 미업로드")
        self.status_correction = QLabel("소급 보정: 0건")
        self.status_api = QLabel("API: 미연결")
        status.addWidget(self.status_excel, 1)
        status.addWidget(self._status_divider())
        status.addWidget(self.status_correction, 1)
        status.addWidget(self._status_divider())
        status.addWidget(self.status_api, 1)
        credit = QLabel(AUTHOR_CREDIT)
        credit.setObjectName("creditLabel")
        status.addPermanentWidget(credit)

    @staticmethod
    def _status_divider() -> QFrame:
        line = QFrame()
        line.setObjectName("statusDivider")
        line.setFrameShape(QFrame.Shape.VLine)
        line.setFixedWidth(1)
        return line

    def _load_api_key(self) -> None:
        key = load_config().get("gemini_api_key", "")
        if key:
            self.api_key_input.setText(key)

    def _save_api_key(self) -> None:
        key = self.api_key_input.text().strip()
        if not key:
            QMessageBox.warning(self, "경고", "API Key를 입력해 주세요.")
            return
        save_config({"gemini_api_key": key})
        QMessageBox.information(self, "저장 완료", "API Key가 config.json에 저장되었습니다.")
        self._test_api_connection()

    def _set_api_status(self, ok: bool, message: str) -> None:
        full = f"API: {message}"
        self.status_api.setText(full if len(full) <= 100 else full[:97] + "...")
        self.status_api.setToolTip(message)
        self.status_api.setStyleSheet(
            "color: #15803d; font-weight: 600;" if ok else "color: #b91c1c; font-weight: 600;"
        )

    def _test_api_connection(self) -> None:
        if self.test_worker is not None and self.test_worker.isRunning():
            self.status_api.setText("API: 연결 확인 중... (이미 진행 중)")
            return
        key = self.api_key_input.text().strip()
        if not key:
            self._set_api_status(False, "키 없음 — API Key를 입력해 주세요.")
            return
        self.btn_test.setEnabled(False)
        self.btn_test.setText("확인 중...")
        self.status_api.setText("API: 연결 확인 중...")
        self.status_api.setStyleSheet("color: #64748b;")
        self.test_worker = GeminiTestWorker(key)
        self.test_worker.finished.connect(self._on_api_test_finished)
        self.test_worker.start()

    def _on_api_test_finished(self, ok: bool, message: str) -> None:
        self.btn_test.setEnabled(True)
        self.btn_test.setText("연결 테스트")
        self._set_api_status(ok, message)

    def _browse_file(self) -> None:
        from PyQt6.QtWidgets import QFileDialog

        path, _ = QFileDialog.getOpenFileName(self, "엑셀 파일 선택", "", "Excel Files (*.xlsx *.xls)")
        if path:
            self._load_excel(path)

    def _load_excel(self, file_path: str) -> None:
        self.status_excel.setText(f"파일: 처리 중... ({Path(file_path).name})")
        self.excel_worker = ExcelWorker(file_path)
        self.excel_worker.finished.connect(self._on_excel_loaded)
        self.excel_worker.error.connect(self._on_excel_error)
        self.excel_worker.start()

    def _on_excel_error(self, message: str) -> None:
        self.status_excel.setText("파일: 오류")
        QMessageBox.critical(self, "파일 처리 오류", message)

    def _on_excel_loaded(self, data: dict[str, Any]) -> None:
        self.inventory_data = data
        changed = sum(1 for it in data["items"] if it["has_change"])
        self.status_excel.setText(f"파일: {data['file_name']} ({data['row_count']}품목)")
        self.status_correction.setText(
            f"소급 보정: {data['correction_count']:,}건 · 변동품목 {changed}건"
        )
        self._populate_table(data)
        self._populate_filters(data)
        self._refresh_3d()
        self._run_ai_analysis(data)
        self.tabs.setCurrentIndex(0)

    def _populate_table(self, data: dict[str, Any]) -> None:
        meta_cols = data["meta_cols"]
        max_pairs = data["max_pair_count"]
        headers = list(meta_cols)
        for i in range(1, max_pairs + 1):
            headers += [f"변경일자{i}", f"재고량{i}(연도말원본)", f"재고량{i}(보정)"]

        self.table.clear()
        self.table.setColumnCount(len(headers))
        self.table.setRowCount(data["row_count"])
        self.table.setHorizontalHeaderLabels(headers)
        highlight = QColor("#dbeafe")

        for row_idx, item in enumerate(data["items"]):
            col = 0
            for meta_val in item["meta"]:
                self.table.setItem(row_idx, col, QTableWidgetItem(str(meta_val)))
                col += 1
            for i in range(max_pairs):
                if i < len(item["dates"]):
                    date_text = item["dates"][i]
                    orig = item["original"][i] if i < len(item["original"]) else None
                    corr = item["corrected"][i]
                    orig_text = "" if orig is None else f"{orig:g}"
                    corr_text = "" if corr is None else f"{corr:g}"
                else:
                    date_text = orig_text = corr_text = ""
                    orig = corr = None
                self.table.setItem(row_idx, col, QTableWidgetItem(date_text))
                self.table.setItem(row_idx, col + 1, QTableWidgetItem(orig_text))
                corr_item = QTableWidgetItem(corr_text)
                if orig is not None and corr is not None and abs(orig - corr) > 1e-9:
                    corr_item.setBackground(highlight)
                self.table.setItem(row_idx, col + 2, corr_item)
                col += 3
            first = self.table.item(row_idx, 0)
            if first is not None:
                first.setData(Qt.ItemDataRole.UserRole, item["label"])

    def _populate_filters(self, data: dict[str, Any]) -> None:
        self._updating_combo = True
        cats = ["전체"] + list(data.get("categories", []))
        for combo in (self.category_combo, self.viz_category_combo):
            combo.blockSignals(True)
            combo.clear()
            combo.addItems(cats)
            combo.setCurrentIndex(0)
            combo.blockSignals(False)
        self._refill_item_combo("전체")
        self._updating_combo = False
        self.chart._show_placeholder("품목 또는 표준품구분을 선택하면 재고 추이 차트가 표시됩니다.")

    def _filtered_items(self, category: str) -> list[dict[str, Any]]:
        if not self.inventory_data:
            return []
        items = self.inventory_data["items"]
        if category == "전체":
            return items
        return [it for it in items if it.get("std_type") == category]

    def _refill_item_combo(self, category: str) -> None:
        self.item_combo.blockSignals(True)
        self.item_combo.clear()
        self._item_by_label.clear()
        self.item_combo.addItem("품목을 선택하세요")
        for item in self._filtered_items(category):
            label = item["label"]
            if label in self._item_by_label:
                label = f"{label} [행{item['row_index'] + 1}]"
                item = {**item, "label": label}
            self._item_by_label[label] = item
            self.item_combo.addItem(label)
        self.item_combo.setCurrentIndex(0)
        self.item_combo.blockSignals(False)

    def _on_category_changed(self, _index: int = 0) -> None:
        if self._updating_combo or not self.inventory_data:
            return
        category = self.category_combo.currentText()
        self._refill_item_combo(category)
        if category != "전체":
            self.chart.plot_category(category, self._filtered_items(category))
        else:
            self.chart._show_placeholder("품목을 선택하거나 표준품구분을 지정해 주세요.")

    def _on_item_combo_changed(self, index: int) -> None:
        if self._updating_combo or index <= 0:
            return
        item = self._item_by_label.get(self.item_combo.currentText().strip())
        if item:
            self.chart.plot_item(
                item["label"],
                item["dates"],
                item["corrected"],
                item.get("original"),
                item.get("original_dates"),
            )

    def _select_item_by_label(self, label: str) -> bool:
        idx = self.item_combo.findText(label)
        if idx < 0:
            for i in range(1, self.item_combo.count()):
                text = self.item_combo.itemText(i)
                if text.startswith(label) or label in text:
                    idx = i
                    break
        if idx < 0:
            return False
        self.item_combo.setCurrentIndex(idx)
        return True

    def _on_table_double_clicked(self, row: int, _column: int) -> None:
        if not self.inventory_data:
            return
        label = None
        first = self.table.item(row, 0)
        if first is not None:
            label = first.data(Qt.ItemDataRole.UserRole)
        if not label and 0 <= row < len(self.inventory_data["items"]):
            label = self.inventory_data["items"][row]["label"]
            std = self.inventory_data["items"][row].get("std_type")
            if std:
                ci = self.category_combo.findText(std)
                if ci >= 0:
                    self.category_combo.setCurrentIndex(ci)
        self.tabs.setCurrentIndex(1)
        if label:
            self._select_item_by_label(str(label))

    def _refresh_3d(self, _index: int = 0) -> None:
        if not self.inventory_data:
            return
        category = self.viz_category_combo.currentText() or "전체"
        stock_items: list[StockItem] = list(self.inventory_data.get("stock_items") or [])
        if category != "전체":
            stock_items = [it for it in stock_items if it.std_type == category]
        try:
            records = build_scatter3d_records(stock_items)
        except Exception as exc:
            self.chart3d.show_message(
                "3D 산점도 데이터를 준비하지 못했습니다.\n"
                "엑셀 형식(등록일자·변경일자·재고량·잔고)을 확인해 주세요.\n\n"
                f"{exc}"
            )
            return
        if not records:
            self.chart3d.show_message(
                "3D 산점도로 표시할 품목이 없습니다.\n"
                "등록일자와 변경이력(변경일자/재고량)이 있는 행이 필요합니다."
            )
            return
        self.chart3d.plot_records(
            records,
            source_file=str(self.inventory_data.get("file_name") or ""),
        )

    def _run_ai_analysis(self, data: dict[str, Any]) -> None:
        key = self.api_key_input.text().strip()
        stock_items = data.get("stock_items") or []
        changed = [it for it in stock_items if it.has_stock_change]
        if not changed:
            self.report_edit.setPlainText(
                "분석 대상 없음\n\n수량 변화가 있는 품목이 없어 AI 분석을 건너뛰었습니다."
            )
            return
        if not key:
            self.report_edit.setPlainText(
                "API Key가 설정되지 않았습니다.\n상단에서 Gemini API Key를 입력하고 [저장]을 눌러 주세요.\n\n"
                f"(변동 품목 {len(changed)}건 대기 중)"
            )
            return
        self.report_edit.setPlainText(
            f"AI 분석 리포트 생성 중... (변동 품목 {len(changed)}건)\n잠시만 기다려 주세요."
        )
        prompt = build_ai_prompt(stock_items)
        self.gemini_worker = GeminiWorker(key, prompt)
        self.gemini_worker.finished.connect(self._on_report_ready)
        self.gemini_worker.error.connect(self._on_report_error)
        self.gemini_worker.start()

    def _on_report_ready(self, text: str) -> None:
        self.report_edit.setMarkdown(text)
        self._set_api_status(True, f"연결됨 ({get_active_gemini_model()})")

    def _on_report_error(self, message: str) -> None:
        self.report_edit.setPlainText(f"AI 분석 생성 실패:\n\n{message}")
        self._set_api_status(False, message)


def create_splash(app: QApplication) -> tuple[QSplashScreen, QProgressBar]:
    """반투명 패널 + 아웃라인 테두리 스플래시."""
    w, h = 540, 320
    pix = QPixmap(w, h)
    pix.fill(Qt.GlobalColor.transparent)

    painter = QPainter(pix)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    painter.setRenderHint(QPainter.RenderHint.TextAntialiasing)

    panel = QPainterPath()
    panel.addRoundedRect(14, 14, w - 28, h - 28, 18, 18)

    # 반투명 유리창 느낌의 배경
    fill = QColor("#f4f7fb")
    fill.setAlpha(210)
    painter.setPen(Qt.PenStyle.NoPen)
    painter.setBrush(QBrush(fill))
    painter.drawPath(panel)

    # 은은한 상단 하이라이트 (반투명)
    highlight = QColor("#ffffff")
    highlight.setAlpha(70)
    painter.setBrush(QBrush(highlight))
    painter.drawRoundedRect(18, 18, w - 36, (h - 36) // 2, 16, 16)

    # 바깥 프레임 테두리
    outer = QPen(QColor("#1e3a5f"))
    outer.setWidth(2)
    outer.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
    painter.setPen(outer)
    painter.setBrush(QBrush(Qt.BrushStyle.NoBrush))
    painter.drawPath(panel)

    # 안쪽 얇은 이중 테두리
    inner = QPen(QColor("#6b87a8"))
    inner.setWidth(1)
    painter.setPen(inner)
    painter.drawRoundedRect(26, 26, w - 52, h - 52, 14, 14)

    # 상단 로고 마크: 마름모+원 아웃라인
    cx, cy = w // 2, 78
    logo_pen = QPen(QColor("#1e3a5f"))
    logo_pen.setWidth(2)
    painter.setPen(logo_pen)
    painter.setBrush(QBrush(Qt.BrushStyle.NoBrush))
    diamond = QPainterPath()
    diamond.moveTo(cx, cy - 22)
    diamond.lineTo(cx + 22, cy)
    diamond.lineTo(cx, cy + 22)
    diamond.lineTo(cx - 22, cy)
    diamond.closeSubpath()
    painter.drawPath(diamond)
    painter.drawEllipse(cx - 10, cy - 10, 20, 20)

    # 타이틀 / 버전
    painter.setPen(QColor("#0b1f3a"))
    painter.setFont(QFont("Malgun Gothic", 16, QFont.Weight.Bold))
    painter.drawText(40, 130, w - 80, 36, Qt.AlignmentFlag.AlignHCenter, "생약표준품 재고 분석 시스템")
    painter.setFont(QFont("Malgun Gothic", 11))
    painter.setPen(QColor("#3d5a80"))
    painter.drawText(40, 168, w - 80, 28, Qt.AlignmentFlag.AlignHCenter, f"{APP_VERSION} 초기화 중...")
    painter.end()

    splash = QSplashScreen(pix)
    splash.setWindowFlag(Qt.WindowType.FramelessWindowHint, True)
    splash.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
    splash.show()

    bar = QProgressBar(splash)
    bar.setGeometry(70, 250, w - 140, 16)
    bar.setRange(0, 100)
    bar.setValue(0)
    bar.setTextVisible(False)
    bar.setStyleSheet(
        """
        QProgressBar {
            background: rgba(255, 255, 255, 120);
            border: 1px solid #1e3a5f;
            border-radius: 8px;
        }
        QProgressBar::chunk {
            background-color: rgba(30, 58, 95, 210);
            border-radius: 7px;
            margin: 2px;
        }
        """
    )
    bar.show()
    app.processEvents()
    return splash, bar


def main() -> None:
    # QWebEngineView 사용 시 QApplication 생성 전에 필요
    try:
        QApplication.setAttribute(Qt.ApplicationAttribute.AA_ShareOpenGLContexts)
    except Exception:
        pass
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(APP_STYLE)

    splash, bar = create_splash(app)
    window = MainWindow()

    steps = [15, 35, 55, 75, 90, 100]
    labels = [
        "모듈 로드 중...",
        "UI 구성 중...",
        "차트 엔진 준비 중...",
        "AI 연동 준비 중...",
        "마무리 중...",
        f"생약표준품 재고 분석 시스템 {APP_VERSION} 초기화 중...",
    ]

    state = {"i": 0}

    def tick() -> None:
        i = state["i"]
        if i < len(steps):
            bar.setValue(steps[i])
            splash.showMessage(
                labels[i],
                Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignHCenter,
                QColor("#3d5a80"),
            )
            state["i"] += 1
            app.processEvents()
            QTimer.singleShot(320, tick)
        else:
            splash.finish(window)
            window.show()

    QTimer.singleShot(200, tick)
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
