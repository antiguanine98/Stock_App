"""
생약표준품 재고 분석 및 소급 보정 시스템 (PyQt6) v1.36
"""

from __future__ import annotations

import json
import sys
import traceback
from pathlib import Path
from typing import Any
from urllib.parse import unquote

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
    QKeySequence,
    QPainter,
    QPainterPath,
    QPen,
    QPixmap,
    QShortcut,
    QTextCharFormat,
    QTextCursor,
    QTextDocument,
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
    QProgressDialog,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSplitter,
    QStatusBar,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextBrowser,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from stock_logic import (
    StockItem,
    build_ai_prompt,
    build_followup_prompt,
    build_scatter3d_records,
    collect_ai_analysis_flags,
    extract_mentioned_codes_from_report,
    format_compendium_context,
    format_compendium_match_report,
    load_compendium_excel,
    lookup_pharmacopoeia_tag,
    markdown_report_to_collapsible_html,
    match_compendium_inventory,
    process_excel,
    process_excels,
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

def _app_dir() -> Path:
    """개발 실행과 PyInstaller 동결 실행 모두에서 리소스 경로를 찾는다."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


def _writable_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


CONFIG_PATH = _writable_dir() / "config.json"
VIEWER_HTML_PATH = _app_dir() / "viewer.html"
APP_VERSION = "v1.36"
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
QLabel#settingsHeaderLabel {
    color: #0b1f3a;
    font-size: 13px;
    font-weight: 700;
}
QPushButton#collapseToggleBtn {
    background-color: #eef3f9;
    color: #1e3a5f;
    border: 1px solid #c5d0de;
    border-radius: 8px;
    padding: 6px 12px;
    font-weight: 600;
}
QPushButton#collapseToggleBtn:hover {
    background-color: #dce6f2;
    border-color: #1e3a5f;
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

    def __init__(self, file_paths: list[str] | str):
        super().__init__()
        if isinstance(file_paths, str):
            self.file_paths = [file_paths]
        else:
            self.file_paths = list(file_paths)

    def run(self) -> None:
        try:
            if len(self.file_paths) == 1:
                self.finished.emit(process_excel(self.file_paths[0]))
            else:
                self.finished.emit(process_excels(self.file_paths))
        except Exception as e:
            self.error.emit(str(e))


class GeminiWorker(QThread):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, api_key: str, prompt: str, *, followup: bool = False):
        super().__init__()
        self.api_key = api_key
        self.prompt = prompt
        self.followup = followup

    def run(self) -> None:
        try:
            self.finished.emit(generate_gemini_report(self.api_key, self.prompt))
        except Exception as e:
            detail = format_gemini_error(e)
            kind = "후속 질문" if self.followup else "AI 분석"
            log_gemini("ERROR", f"{kind} 실패: {detail}", e)
            self.error.emit(detail)


class CompendiumWorker(QThread):
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, file_path: str):
        super().__init__()
        self.file_path = file_path

    def run(self) -> None:
        try:
            self.finished.emit(load_compendium_excel(self.file_path))
        except Exception as e:
            self.error.emit(str(e))


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
    files_dropped = pyqtSignal(list)

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        title: str = "엑셀 파일을 여기에 드래그 앤 드롭 (다중 가능)",
        subtitle: str = ".xlsx / .xls · 2개 이상 동시/순차 업로드 시 통합 분석 · [파일 선택] 지원",
        min_height: int = 96,
    ):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setMinimumHeight(min_height)
        self.setObjectName("dropZone")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 14)
        title_lbl = QLabel(title)
        title_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_lbl.setWordWrap(True)
        title_lbl.setStyleSheet(
            "color: #1e3a5f; font-size: 14px; font-weight: 700; border: none; background: transparent;"
        )
        subtitle_lbl = QLabel(subtitle)
        subtitle_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle_lbl.setWordWrap(True)
        subtitle_lbl.setStyleSheet(
            "color: #5b6b7c; font-size: 12px; border: none; background: transparent;"
        )
        layout.addStretch()
        layout.addWidget(title_lbl)
        layout.addWidget(subtitle_lbl)
        layout.addStretch()

    @staticmethod
    def _excel_paths_from_urls(urls) -> list[str]:
        paths: list[str] = []
        for url in urls:
            path = url.toLocalFile()
            if path.lower().endswith((".xlsx", ".xls")):
                paths.append(path)
        return paths

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls() and self._excel_paths_from_urls(event.mimeData().urls()):
            event.acceptProposedAction()
            return
        event.ignore()

    def dropEvent(self, event: QDropEvent) -> None:
        paths = self._excel_paths_from_urls(event.mimeData().urls())
        if paths:
            self.files_dropped.emit(paths)
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
        if self._cursor is not None:
            try:
                self._cursor.remove()
            except Exception:
                pass
            self._cursor = None
        self._annot = None
        self._points = []

    def _show_placeholder(self, message: str) -> None:
        self._clear_helpers()
        self.figure.clf()
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
        xs = [p[0] for p in self._points]
        ys = [p[1] for p in self._points if p[1] == p[1]]
        x_span = max(1.0, max(xs) - min(xs))
        y_span = max(1.0, (max(ys) - min(ys)) if ys else 1.0)
        nd = ((nearest[0] - event.xdata) / x_span) ** 2 + (
            (nearest[1] - event.ydata) / y_span
        ) ** 2
        if nd < 0.04:
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
        pharmacopoeia_tag: str = "",
        subtitle: str = "",
    ) -> None:
        if not dates:
            self._show_placeholder(f"{label}\n표시할 변동기록이 없습니다.")
            return

        self._clear_helpers()
        self.figure.clf()
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        ax.set_facecolor("#fafbfc")

        def _years(vals: list[str]) -> list[int]:
            out: list[int] = []
            for d in vals:
                try:
                    out.append(int(str(d)[:4]))
                except (TypeError, ValueError):
                    out.append(len(out))
            return out

        years = _years(dates)
        y_corr = [v if v is not None else float("nan") for v in corrected]

        if original is not None and original_dates is not None and original_dates == dates:
            y_orig = [v if v is not None else float("nan") for v in original]
            ax.plot(years, y_orig, "o--", label="연도말 원본", color="#c0392b", linewidth=1.5, markersize=5, alpha=0.7)

        (line,) = ax.plot(
            years, y_corr, "s-", label="연도말 소급 보정",
            color="#1e3a5f", linewidth=2.2, markersize=7,
        )
        labels = [
            f"{label}\n연도: {y}\n재고량: {q:g}" if q == q else f"{label}\n연도: {y}"
            for y, q in zip(years, y_corr)
        ]
        self._attach_hover(ax, line, labels)

        ax.set_xticks(years)
        ax.set_xticklabels([str(y) for y in years], rotation=0, ha="center", fontsize=9)
        ax.set_xlabel("연도 (YYYY)")
        ax.set_ylabel("재고량")
        tag = (pharmacopoeia_tag or subtitle or "").strip()
        title = f"재고 추이 — {label}"
        if tag:
            title = f"{title}  {tag}"
        ax.set_title(title, fontsize=12, fontweight="bold", color="#0b1f3a", pad=12)
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
        self.figure.clf()
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        ax.set_facecolor("#fafbfc")
        palette = ["#1e3a5f", "#2563eb", "#0f766e", "#b45309", "#be123c", "#7c3aed", "#0369a1"]
        hover_labels: list[str] = []
        lines = []
        all_years: set[int] = set()

        for i, item in enumerate(valid[:12]):
            dates = item["dates"]
            try:
                years = [int(str(d)[:4]) for d in dates]
            except (TypeError, ValueError):
                years = list(range(len(dates)))
            all_years.update(years)
            y = [v if v is not None else float("nan") for v in item["corrected"]]
            (line,) = ax.plot(
                years, y, "o-",
                label=item["label"][:28],
                color=palette[i % len(palette)],
                linewidth=1.8,
                markersize=5,
            )
            lines.append(line)
            for yr, q in zip(years, y):
                hover_labels.append(
                    f"{item['label']}\n연도: {yr}\n재고량: {q:g}" if q == q else f"{item['label']}\n연도: {yr}"
                )

        if lines:
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
                            f"{item['label']}\n연도: {item['dates'][di]}\n재고량: {qtxt}"
                        )
                    sel.annotation.get_bbox_patch().set(fc="#0b1f3a", alpha=0.92)
                    sel.annotation.set_color("white")
            else:
                self._attach_hover(ax, lines[0], hover_labels)

        if all_years:
            ticks = sorted(all_years)
            ax.set_xticks(ticks)
            ax.set_xticklabels([str(y) for y in ticks], fontsize=8)
        ax.set_xlabel("연도 (YYYY)")
        ax.set_ylabel("재고량")
        ax.set_title(f"카테고리 비교 — {category}", fontsize=12, fontweight="bold", color="#0b1f3a")
        ax.legend(loc="best", fontsize=8, frameon=False)
        ax.grid(True, axis="y", alpha=0.28, linestyle="--")
        self.figure.tight_layout()
        self.draw()


class Scatter3DView(QWidget):
    """viewer.html 기반 3D 산점도 (QWebEngineView)."""

    point_picked = pyqtSignal(str, str)  # code, name

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setObjectName("scatter3dHost")
        self.setStyleSheet(
            "#scatter3dHost { background-color: #f7faf8; border: 1px solid #d5e0da; border-radius: 8px; }"
        )
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        self._clearing_hash = False

        self._fallback = QLabel(
            "엑셀을 로드하면 AI 추이 기준(연평균 분양량 · 예상 소진기간 · 순감소량) 3D 산점도가 표시됩니다."
        )
        self._fallback.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._fallback.setWordWrap(True)
        self._fallback.setStyleSheet("color: #3d5348; font-size: 13px; padding: 24px;")
        layout.addWidget(self._fallback)

        self._web: Any = None
        if QWebEngineView is not None:
            try:
                web = QWebEngineView(self)
                web.setContextMenuPolicy(Qt.ContextMenuPolicy.NoContextMenu)
                web.page().setBackgroundColor(QColor(247, 250, 248))
                web.urlChanged.connect(self._on_url_changed)
                layout.addWidget(web, stretch=1)
                self._web = web
                self._fallback.hide()
                self.show_message(
                    "엑셀을 로드하면 AI 추이 기준(연평균 분양량 · 예상 소진기간 · 순감소량) 3D 산점도가 표시됩니다."
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

    def _on_url_changed(self, url: QUrl) -> None:
        if self._clearing_hash:
            return
        frag = (url.fragment() or "").strip()
        if not frag.startswith("pick:"):
            return
        payload = frag[5:]
        parts = payload.split("|", 1)
        if len(parts) != 2:
            return
        code = unquote(parts[0])
        name = unquote(parts[1])
        self._clearing_hash = True
        try:
            if self._web is not None:
                self._web.page().runJavaScript(
                    "try{history.replaceState(null,'',location.pathname+location.search);}catch(e){}"
                )
        finally:
            QTimer.singleShot(0, self._reset_clearing_hash)
        self.point_picked.emit(code, name)

    def _reset_clearing_hash(self) -> None:
        self._clearing_hash = False

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
html,body{{margin:0;height:100%;background:#f7faf8;color:#3d5348;
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
        self._loaded_paths: list[str] = []
        self._chat_history: list[dict[str, str]] = []
        self._initial_report: str = ""
        self._report_expanded_ids: set[str] = set()
        self._chat_busy = False
        self.compendium_df = None
        self.compendium_meta: dict[str, Any] | None = None
        self.compendium_worker: CompendiumWorker | None = None
        self._busy_dialog: QProgressDialog | None = None

        self._build_ui()
        self._load_api_key()

    def _show_busy(self, title: str, text: str) -> None:
        """작업 중 로딩창. Windows에서 바로 보이도록 즉시 표시·최상위 유지."""
        self._close_busy()
        dlg = QProgressDialog(text, None, 0, 0, self)
        dlg.setWindowTitle(title)
        dlg.setWindowModality(Qt.WindowModality.ApplicationModal)
        dlg.setMinimumDuration(0)
        dlg.setCancelButton(None)
        dlg.setAutoClose(False)
        dlg.setAutoReset(False)
        dlg.setMinimumWidth(380)
        dlg.setWindowFlags(
            dlg.windowFlags()
            | Qt.WindowType.CustomizeWindowHint
            | Qt.WindowType.WindowTitleHint
            | Qt.WindowType.WindowStaysOnTopHint
        )
        self._busy_dialog = dlg
        dlg.show()
        dlg.raise_()
        dlg.activateWindow()
        # 메인 스레드 블로킹 전에 반드시 한 번 페인트
        QApplication.processEvents()

    def _close_busy(self) -> None:
        if self._busy_dialog is not None:
            dlg = self._busy_dialog
            self._busy_dialog = None
            dlg.hide()
            dlg.close()
            dlg.deleteLater()
            QApplication.processEvents()

    def _run_after_busy_paint(self, fn) -> None:
        """로딩창이 먼저 그려진 뒤 무거운 동기 작업/워커 시작."""
        QTimer.singleShot(40, fn)

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setSpacing(10)
        root.setContentsMargins(18, 12, 18, 8)

        title = QLabel(f"생약표준품 재고 분석 시스템 {APP_VERSION}")
        title.setObjectName("titleLabel")
        title.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        root.addWidget(title, stretch=0)

        # --- 접이식 상단: 엑셀 / 공정서 / API Key ---
        self._settings_collapsed = False
        settings_card = QFrame()
        settings_card.setObjectName("card")
        settings_card.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum)
        self._settings_card = settings_card
        settings_outer = QVBoxLayout(settings_card)
        settings_outer.setContentsMargins(14, 10, 14, 10)
        settings_outer.setSpacing(8)

        header_row = QHBoxLayout()
        header_row.setSpacing(8)
        self._settings_header_label = QLabel("설정 및 파일 업로드")
        self._settings_header_label.setObjectName("settingsHeaderLabel")
        header_row.addWidget(self._settings_header_label, stretch=1)
        self.btn_toggle_settings = QPushButton("▲ 설정 및 파일 업로드 영역 접기")
        self.btn_toggle_settings.setObjectName("collapseToggleBtn")
        self.btn_toggle_settings.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_toggle_settings.setMinimumHeight(34)
        self.btn_toggle_settings.clicked.connect(self._toggle_settings_panel)
        header_row.addWidget(self.btn_toggle_settings, stretch=0)
        settings_outer.addLayout(header_row)

        self._settings_body = QWidget()
        self._settings_body.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum
        )
        top_layout = QVBoxLayout(self._settings_body)
        top_layout.setContentsMargins(0, 2, 0, 0)
        top_layout.setSpacing(10)

        file_label = QLabel("재고 엑셀 업로드")
        file_label.setObjectName("sectionLabel")
        top_layout.addWidget(file_label)

        file_row = QHBoxLayout()
        self.dropzone = DropZone(min_height=78)
        self.dropzone.files_dropped.connect(self._load_excels)
        file_row.addWidget(self.dropzone, stretch=1)
        btn_col = QVBoxLayout()
        btn_select = QPushButton("파일 선택")
        btn_select.setObjectName("secondaryBtn")
        btn_select.setMinimumWidth(110)
        btn_select.setMinimumHeight(42)
        btn_select.clicked.connect(self._browse_file)
        btn_clear = QPushButton("목록 초기화")
        btn_clear.setObjectName("secondaryBtn")
        btn_clear.setMinimumWidth(110)
        btn_clear.clicked.connect(self._clear_loaded_files)
        btn_col.addWidget(btn_select)
        btn_col.addWidget(btn_clear)
        file_row.addLayout(btn_col)
        top_layout.addLayout(file_row)

        comp_label = QLabel("공정서 DB 업로드 (규격 참조 · 재고 분석 제외)")
        comp_label.setObjectName("sectionLabel")
        top_layout.addWidget(comp_label)

        comp_row = QHBoxLayout()
        self.compendium_dropzone = DropZone(
            title="공정서 DB 엑셀을 드래그 앤 드롭",
            subtitle=".xlsx / .xls · 재고 시계열로 파싱하지 않음 · AI 규격/기준 참조 전용",
            min_height=64,
        )
        self.compendium_dropzone.files_dropped.connect(self._on_compendium_dropped)
        comp_row.addWidget(self.compendium_dropzone, stretch=1)
        comp_btn_col = QVBoxLayout()
        btn_comp_select = QPushButton("공정서 선택")
        btn_comp_select.setObjectName("secondaryBtn")
        btn_comp_select.setMinimumWidth(110)
        btn_comp_select.clicked.connect(self._browse_compendium)
        btn_comp_clear = QPushButton("공정서 해제")
        btn_comp_clear.setObjectName("secondaryBtn")
        btn_comp_clear.setMinimumWidth(110)
        btn_comp_clear.clicked.connect(self._clear_compendium)
        comp_btn_col.addWidget(btn_comp_select)
        comp_btn_col.addWidget(btn_comp_clear)
        comp_row.addLayout(comp_btn_col)
        top_layout.addLayout(comp_row)
        self.compendium_status = QLabel("공정서 DB: 미등록")
        self.compendium_status.setStyleSheet("color: #64748b; font-size: 12px;")
        top_layout.addWidget(self.compendium_status)

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

        settings_outer.addWidget(self._settings_body)
        root.addWidget(settings_card, stretch=0)

        self.tabs = QTabWidget()
        self.tabs.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.tabs.setMinimumHeight(280)

        # Tab 1
        table_wrap = QWidget()
        table_layout = QVBoxLayout(table_wrap)
        table_layout.setContentsMargins(8, 8, 8, 8)
        hint = QLabel(
            "행을 더블클릭하면 해당 품목의 재고 추이 차트로 이동합니다. "
            "(등록일자 제외 · 잔고→재고 · 변경일자는 연도 YYYY)"
        )
        hint.setStyleSheet("color: #64748b; font-size: 12px;")
        table_layout.addWidget(hint)
        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(False)
        self.table.cellDoubleClicked.connect(self._on_table_double_clicked)
        table_layout.addWidget(self.table, stretch=1)
        self.tabs.addTab(table_wrap, "보정 데이터 표")

        # Tab 2 chart
        chart_widget = QWidget()
        chart_layout = QVBoxLayout(chart_widget)
        chart_layout.setContentsMargins(10, 10, 10, 10)
        chart_layout.setSpacing(8)

        filter_row = QHBoxLayout()
        # category_combo는 표 더블클릭/_populate_filters/_refill_item_combo용으로 유지 (차트 UI에는 미표시)
        self.category_combo = QComboBox()
        self.category_combo.addItem("전체")
        self.category_combo.setMinimumWidth(160)
        self.category_combo.currentIndexChanged.connect(self._on_category_changed)
        self.category_combo.hide()

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

        hover_hint = QLabel("그래프 포인트에 마우스를 올리면 연도·재고량이 표시됩니다. (X축: 연도 YYYY, 빈 연도 보간)")
        hover_hint.setStyleSheet("color: #64748b; font-size: 12px;")
        chart_layout.addWidget(hover_hint)

        self.chart = InventoryChart()
        chart_layout.addWidget(self.chart, stretch=1)
        self.tabs.addTab(chart_widget, "재고 추이 차트")

        # Tab 3 AI chatbot — left: fixed report / right: interactive chat
        report_wrap = QWidget()
        report_layout = QVBoxLayout(report_wrap)
        report_layout.setContentsMargins(8, 8, 8, 8)
        report_layout.setSpacing(8)
        chat_hint = QLabel(
            "초기 분석 리포트 생성 후 추가 질문을 입력하세요. "
            "후속 질문은 재고 통합 데이터·공정서 DB를 실시간 재검토하여 답변합니다."
        )
        chat_hint.setStyleSheet("color: #64748b; font-size: 12px;")
        report_layout.addWidget(chat_hint)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)

        left_pane = QWidget()
        left_layout = QVBoxLayout(left_pane)
        left_layout.setContentsMargins(0, 0, 4, 0)
        left_layout.setSpacing(6)
        left_title = QLabel("표준 분석 리포트")
        left_title.setObjectName("sectionLabel")
        left_layout.addWidget(left_title)

        # Ctrl+F 찾기 바 (기본 숨김)
        self.report_find_bar = QWidget()
        find_row = QHBoxLayout(self.report_find_bar)
        find_row.setContentsMargins(0, 0, 0, 0)
        find_row.setSpacing(6)
        self.report_find_input = QLineEdit()
        self.report_find_input.setPlaceholderText("리포트에서 찾기…")
        self.report_find_input.returnPressed.connect(lambda: self._find_in_report(forward=True))
        find_row.addWidget(self.report_find_input, stretch=1)
        self.btn_find_prev = QPushButton("이전")
        self.btn_find_prev.clicked.connect(lambda: self._find_in_report(forward=False))
        find_row.addWidget(self.btn_find_prev)
        self.btn_find_next = QPushButton("다음")
        self.btn_find_next.clicked.connect(lambda: self._find_in_report(forward=True))
        find_row.addWidget(self.btn_find_next)
        self.btn_find_close = QPushButton("닫기")
        self.btn_find_close.clicked.connect(self._hide_report_find_bar)
        find_row.addWidget(self.btn_find_close)
        self.report_find_bar.hide()
        left_layout.addWidget(self.report_find_bar)

        self.report_fixed = QTextBrowser()
        self.report_fixed.setReadOnly(True)
        self.report_fixed.setOpenExternalLinks(False)
        self.report_fixed.setOpenLinks(False)
        self.report_fixed.anchorClicked.connect(self._on_report_anchor_clicked)
        self.report_fixed.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        self.report_fixed.setPlaceholderText(
            "엑셀 업로드 후 생약표준품 분양·소진 예측 AI 리포트가 여기에 표시됩니다."
        )
        self.report_fixed.setFont(QFont("Malgun Gothic", 10))
        self.report_fixed.setStyleSheet(
            "QTextBrowser { padding: 12px; line-height: 1.65; font-family: 'Malgun Gothic'; font-size: 10.5pt; }"
        )
        left_layout.addWidget(self.report_fixed, stretch=1)

        self._report_find_shortcut = QShortcut(QKeySequence.StandardKey.Find, self.report_fixed)
        self._report_find_shortcut.activated.connect(self._show_report_find_bar)
        self._report_find_esc = QShortcut(QKeySequence(Qt.Key.Key_Escape), self.report_find_bar)
        self._report_find_esc.activated.connect(self._hide_report_find_bar)
        splitter.addWidget(left_pane)

        right_pane = QWidget()
        right_layout = QVBoxLayout(right_pane)
        right_layout.setContentsMargins(4, 0, 0, 0)
        right_layout.setSpacing(6)
        right_title = QLabel("대화형 챗봇")
        right_title.setObjectName("sectionLabel")
        right_layout.addWidget(right_title)
        self.chat_view = QTextEdit()
        self.chat_view.setReadOnly(True)
        self.chat_view.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        self.chat_view.setPlaceholderText(
            "추가 질문과 AI 답변이 여기에 표시됩니다."
        )
        self.chat_view.setFont(QFont("Malgun Gothic", 10))
        self.chat_view.setStyleSheet(
            "QTextEdit { padding: 12px; line-height: 1.65; font-family: 'Malgun Gothic'; font-size: 10.5pt; }"
        )
        right_layout.addWidget(self.chat_view, stretch=1)

        chat_row = QHBoxLayout()
        self.chat_input = QLineEdit()
        self.chat_input.setPlaceholderText("추가 질문 또는 세부 분석 요청을 입력하세요…")
        self.chat_input.returnPressed.connect(self._send_chat_message)
        chat_row.addWidget(self.chat_input, stretch=1)
        self.btn_chat_send = QPushButton("전송")
        self.btn_chat_send.setEnabled(False)
        self.btn_chat_send.clicked.connect(self._send_chat_message)
        chat_row.addWidget(self.btn_chat_send)
        self.btn_chat_clear = QPushButton("대화 초기화")
        self.btn_chat_clear.setEnabled(False)
        self.btn_chat_clear.clicked.connect(self._clear_chat_history)
        chat_row.addWidget(self.btn_chat_clear)
        right_layout.addLayout(chat_row)
        splitter.addWidget(right_pane)

        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([560, 560])
        left_pane.setMinimumWidth(280)
        right_pane.setMinimumWidth(280)
        report_layout.addWidget(splitter, stretch=1)
        self.tabs.addTab(report_wrap, "AI 분석 리포트")

        # Tab 4 3D scatter (viewer.html)
        viz_wrap = QWidget()
        viz_layout = QVBoxLayout(viz_wrap)
        viz_layout.setContentsMargins(10, 10, 10, 10)
        viz_row = QHBoxLayout()
        viz_label = QLabel("표준품구분")
        viz_label.setObjectName("sectionLabel")
        viz_row.addWidget(viz_label)
        self.viz_category_combo = QComboBox()
        self.viz_category_combo.addItem("전체")
        self.viz_category_combo.currentIndexChanged.connect(self._refresh_3d)
        viz_row.addWidget(self.viz_category_combo)

        ai_label = QLabel("AI 연동")
        ai_label.setObjectName("sectionLabel")
        viz_row.addWidget(ai_label)
        self.viz_ai_filter = QComboBox()
        self.viz_ai_filter.addItem("전체 표시", "all")
        self.viz_ai_filter.addItem("5년 소진 위험", "deplete")
        self.viz_ai_filter.addItem("최근 분양 급증", "surge")
        self.viz_ai_filter.addItem("소진∩급증 (핵심 위험)", "deplete_surge")
        self.viz_ai_filter.addItem("AI 리포트 언급", "mentioned")
        self.viz_ai_filter.setMinimumWidth(180)
        self.viz_ai_filter.currentIndexChanged.connect(self._refresh_3d)
        viz_row.addWidget(self.viz_ai_filter)
        viz_row.addStretch(1)
        viz_layout.addLayout(viz_row)
        viz_hint = QLabel(
            "클릭 시 재고 추이 차트로 이동 · 드래그 회전 · 휠 확대"
        )
        viz_hint.setStyleSheet("color: #64748b; font-size: 12px;")
        viz_layout.addWidget(viz_hint)
        self.chart3d = Scatter3DView()
        self.chart3d.point_picked.connect(self._on_3d_point_picked)
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

    def _toggle_settings_panel(self) -> None:
        self._set_settings_collapsed(not self._settings_collapsed)

    def _set_settings_collapsed(self, collapsed: bool) -> None:
        self._settings_collapsed = collapsed
        self._settings_body.setVisible(not collapsed)
        if collapsed:
            self.btn_toggle_settings.setText("▼ 설정 및 파일 업로드 영역 펼치기")
            self._settings_header_label.setText("설정 및 파일 업로드 (접힘)")
            # 접힌 뒤에는 헤더 한 줄만 남기고 탭이 세로 공간을 전부 가져감
            self._settings_card.setMaximumHeight(70)
            self._settings_card.setSizePolicy(
                QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
            )
        else:
            self.btn_toggle_settings.setText("▲ 설정 및 파일 업로드 영역 접기")
            self._settings_header_label.setText("설정 및 파일 업로드")
            self._settings_card.setMaximumHeight(16777215)  # QWIDGETSIZE_MAX
            self._settings_card.setSizePolicy(
                QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum
            )
        self.tabs.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self._settings_card.updateGeometry()
        self.tabs.updateGeometry()
        central = self.centralWidget()
        if central is not None and central.layout() is not None:
            central.layout().activate()
        self.update()

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

    def _browse_compendium(self) -> None:
        from PyQt6.QtWidgets import QFileDialog

        path, _ = QFileDialog.getOpenFileName(
            self, "공정서 DB 엑셀 선택", "", "Excel Files (*.xlsx *.xls)"
        )
        if path:
            self._load_compendium(path)

    def _on_compendium_dropped(self, paths: list[str]) -> None:
        if paths:
            self._load_compendium(paths[0])

    def _clear_compendium(self) -> None:
        self.compendium_df = None
        self.compendium_meta = None
        self.compendium_status.setText("공정서 DB: 미등록")
        self.compendium_status.setStyleSheet("color: #64748b; font-size: 12px;")
        self.compendium_status.setToolTip("")
        if self.inventory_data is not None:
            self.inventory_data["compendium_match"] = {}
            self.inventory_data["compendium_match_report"] = ""

    def _load_compendium(self, file_path: str) -> None:
        self.compendium_status.setText(f"공정서 DB: 로딩 중... ({Path(file_path).name})")
        self.compendium_status.setStyleSheet("color: #64748b; font-size: 12px;")
        self._show_busy("공정서 DB", f"공정서 DB를 불러오는 중...\n{Path(file_path).name}")

        def _start() -> None:
            self.compendium_worker = CompendiumWorker(file_path)
            self.compendium_worker.finished.connect(self._on_compendium_loaded)
            self.compendium_worker.error.connect(self._on_compendium_error)
            self.compendium_worker.start()

        self._run_after_busy_paint(_start)

    def _on_compendium_loaded(self, data: dict[str, Any]) -> None:
        self._close_busy()
        self.compendium_df = data.get("dataframe")
        self.compendium_meta = {
            "file_path": data.get("file_path"),
            "file_name": data.get("file_name"),
            "row_count": data.get("row_count"),
            "columns": data.get("columns"),
            "sheet_count": data.get("sheet_count"),
            "entries": data.get("entries") or [],
            "parsed_count": data.get("parsed_count"),
        }
        name = data.get("file_name", "")
        rows = data.get("row_count", 0)
        sheets = data.get("sheet_count", 1)
        parsed = data.get("parsed_count") or len(data.get("entries") or [])
        match_info = self._refresh_compendium_match()
        match_n = (match_info or {}).get("correction_count", 0)
        status = f"공정서 DB: {name} ({rows:,}행 · 시트 {sheets} · 파싱 {parsed})"
        if match_info is not None:
            status += f" · 매칭 {match_n}건"
        self.compendium_status.setText(status)
        self.compendium_status.setStyleSheet("color: #15803d; font-size: 12px; font-weight: 600;")
        self.compendium_status.setToolTip(
            f"{name}\n열: {', '.join((data.get('columns') or [])[:30])}"
        )
        # 이미 재고가 로드된 상태면 다음 AI 질문부터 반영 (초기 리포트는 재실행하지 않음)
        if self._initial_report:
            QMessageBox.information(
                self,
                "공정서 DB 등록",
                "공정서 DB가 등록되었습니다.\n"
                "이후 챗봇 답변과, 재고 파일을 다시 불러와 생성하는 초기 리포트에 규격 참조로 사용됩니다.",
            )

    def _refresh_compendium_match(self) -> dict[str, Any] | None:
        """재고·공정서 모두 있을 때 매칭 결과를 inventory_data에 저장."""
        if not self.inventory_data:
            return None
        entries = []
        if self.compendium_meta:
            entries = list(self.compendium_meta.get("entries") or [])
        stock_items = list(self.inventory_data.get("stock_items") or [])
        if not entries or not stock_items:
            self.inventory_data["compendium_match"] = {}
            self.inventory_data["compendium_match_report"] = ""
            return None
        match = match_compendium_inventory(entries, stock_items)
        self.inventory_data["compendium_match"] = match
        self.inventory_data["compendium_match_report"] = format_compendium_match_report(match)
        return match

    def _pharmacopoeia_tag_for_item(self, item: dict[str, Any] | StockItem | str) -> str:
        match = (self.inventory_data or {}).get("compendium_match") or {}
        if not match:
            return ""
        if isinstance(item, StockItem):
            return lookup_pharmacopoeia_tag(item, match)
        if isinstance(item, str):
            label = item
            mgmt = ""
        else:
            label = str(item.get("label") or "")
            mgmt = str(item.get("mgmt_no") or item.get("manage_no") or "")
        by_label = match.get("by_label") or {}
        if label in by_label:
            return by_label[label]
        for it in (self.inventory_data or {}).get("stock_items") or []:
            if not isinstance(it, StockItem):
                continue
            if (mgmt and it.manage_no == mgmt) or (label and it.label == label):
                return lookup_pharmacopoeia_tag(it, match)
            if label and (it.label in label or label.startswith(it.label)):
                return lookup_pharmacopoeia_tag(it, match)
        return ""

    def _on_compendium_error(self, message: str) -> None:
        self._close_busy()
        self.compendium_status.setText("공정서 DB: 오류")
        self.compendium_status.setStyleSheet("color: #b91c1c; font-size: 12px; font-weight: 600;")
        QMessageBox.critical(self, "공정서 DB 오류", message)

    def _compendium_prompt_context(self, stock_items: list | None = None) -> str | None:
        if self.compendium_df is None:
            return None
        items = stock_items
        if items is None and self.inventory_data:
            items = self.inventory_data.get("stock_items")
        text = format_compendium_context(
            self.compendium_df,
            items or [],
            meta=self.compendium_meta,
        )
        return text or None

    def _browse_file(self) -> None:
        from PyQt6.QtWidgets import QFileDialog

        paths, _ = QFileDialog.getOpenFileNames(
            self, "엑셀 파일 선택 (다중 가능)", "", "Excel Files (*.xlsx *.xls)"
        )
        if paths:
            self._load_excels(list(paths))

    def _clear_loaded_files(self) -> None:
        self._close_busy()
        self._loaded_paths = []
        self.inventory_data = None
        self._chat_history.clear()
        self._initial_report = ""
        self._report_expanded_ids.clear()
        self._set_chat_enabled(False)
        self.report_fixed.clear()
        self.chat_view.clear()
        self.table.clear()
        self.table.setRowCount(0)
        self.table.setColumnCount(0)
        self.status_excel.setText("파일: 미업로드")
        self.status_correction.setText("소급 보정: 0건")
        self.chart._show_placeholder("품목 또는 표준품구분을 선택하면 재고 추이 차트가 표시됩니다.")
        self.chart3d.show_message("엑셀을 다시 업로드해 주세요.")
        self._set_settings_collapsed(False)

    def _load_excel(self, file_path: str) -> None:
        self._load_excels([file_path])

    def _load_excels(self, file_paths: list[str]) -> None:
        if not file_paths:
            return
        # 순차 업로드: 기존 목록에 합치고 중복 제거
        combined = list(dict.fromkeys(self._loaded_paths + list(file_paths)))
        self._loaded_paths = combined
        names = [Path(p).name for p in combined]
        label = names[0] if len(names) == 1 else f"{len(names)}개 파일"
        self.status_excel.setText(f"파일: 처리 중... ({label})")
        self._show_busy("재고 엑셀 처리", f"재고 데이터를 통합·보정하는 중...\n{label}")

        def _start() -> None:
            self.excel_worker = ExcelWorker(combined)
            self.excel_worker.finished.connect(self._on_excel_loaded)
            self.excel_worker.error.connect(self._on_excel_error)
            self.excel_worker.start()

        self._run_after_busy_paint(_start)

    def _on_excel_error(self, message: str) -> None:
        self._close_busy()
        self.status_excel.setText("파일: 오류")
        QMessageBox.critical(self, "파일 처리 오류", message)

    def _on_excel_loaded(self, data: dict[str, Any]) -> None:
        self._close_busy()
        stock_items = list(data.get("stock_items") or [])
        data["ai_flags"] = collect_ai_analysis_flags(stock_items)
        data["ai_mentioned_codes"] = []
        self.inventory_data = data
        self._refresh_compendium_match()
        self._chat_history.clear()
        self._initial_report = ""
        self._report_expanded_ids.clear()
        self._set_chat_enabled(False)
        self.report_fixed.clear()
        self.chat_view.clear()
        changed = sum(1 for it in data["items"] if it["has_change"])
        deplete_n = len(data["ai_flags"].get("deplete_codes") or [])
        surge_n = len(data["ai_flags"].get("surge_codes") or [])
        dash = (data["ai_flags"].get("dashboard") or {}).get("kpis") or []
        val_kpi = next((k for k in dash if k.get("key") == "total_value"), None)
        val_txt = f" · 환산 {val_kpi['display']}" if val_kpi else ""
        names = data.get("file_names") or [data.get("file_name", "")]
        if len(names) <= 2:
            file_label = ", ".join(names)
        else:
            file_label = f"{names[0]} 외 {len(names) - 1}개"
        match = data.get("compendium_match") or {}
        match_n = match.get("correction_count")
        match_txt = f" · 공정서매칭 {match_n}건" if match_n is not None and match else ""
        self.status_excel.setText(f"파일: {file_label} ({data['row_count']}품목)")
        self.status_excel.setToolTip("\n".join(names))
        self.status_correction.setText(
            f"소급 보정: {data['correction_count']:,}건 · 변동 {changed}건 · "
            f"소진후보 {deplete_n} · 가속 {surge_n}{val_txt}{match_txt}"
        )
        self._populate_table(data)
        self._populate_filters(data)
        self._refresh_3d()
        self._run_ai_analysis(data)
        self.tabs.setCurrentIndex(0)
        # 로드 완료 후 상단을 접어 탭/챗봇 영역을 넓힘
        self._set_settings_collapsed(True)

    def _populate_table(self, data: dict[str, Any]) -> None:
        meta_cols = data["meta_cols"]
        max_pairs = data["max_pair_count"]
        headers = list(meta_cols)
        for i in range(1, max_pairs + 1):
            headers += [f"연도{i}", f"재고량{i}(연도말원본)", f"재고량{i}(보정)"]

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
                    date_text = str(item["dates"][i])
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
            tag = self._pharmacopoeia_tag_for_item(item)
            self.chart.plot_item(
                item["label"],
                item["dates"],
                item["corrected"],
                item.get("original"),
                item.get("original_dates"),
                pharmacopoeia_tag=tag,
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

    def _on_3d_point_picked(self, code: str, name: str) -> None:
        if not self.inventory_data:
            return
        items = list(self.inventory_data.get("items") or [])
        target = None
        code = (code or "").strip()
        name = (name or "").strip()
        if code:
            for it in items:
                mgmt = str(it.get("mgmt_no") or it.get("manage_no") or "")
                label = str(it.get("label") or "")
                if mgmt == code or label.startswith(f"{code}(") or code in label:
                    target = it
                    break
        if target is None and name:
            for it in items:
                label = str(it.get("label") or "")
                if name in label or label.endswith(f"({name})"):
                    target = it
                    break
        self.tabs.setCurrentIndex(1)
        if target is None:
            return
        # 품목 목록이 전체 기준으로 보이도록 카테고리 리셋
        if self.category_combo.currentIndex() != 0:
            self._updating_combo = True
            self.category_combo.setCurrentIndex(0)
            self._refill_item_combo("전체")
            self._updating_combo = False
        self._select_item_by_label(str(target["label"]))

    def _refresh_3d(self, _index: int = 0) -> None:
        if not self.inventory_data:
            return
        category = self.viz_category_combo.currentText() or "전체"
        ai_mode = self.viz_ai_filter.currentData() if hasattr(self, "viz_ai_filter") else "all"
        stock_items: list[StockItem] = list(self.inventory_data.get("stock_items") or [])
        if category != "전체":
            stock_items = [it for it in stock_items if it.std_type == category]
        ai_flags = self.inventory_data.get("ai_flags")
        mentioned = set(self.inventory_data.get("ai_mentioned_codes") or [])
        try:
            records = build_scatter3d_records(
                stock_items,
                ai_flags=ai_flags,
                mentioned_codes=mentioned,
            )
        except Exception as exc:
            self.chart3d.show_message(
                "3D 산점도 데이터를 준비하지 못했습니다.\n"
                "감소 추이가 있는 품목(보정 이력)을 확인해 주세요.\n\n"
                f"{exc}"
            )
            return

        if ai_mode == "deplete":
            records = [r for r in records if r.get("depleteWithin5y")]
        elif ai_mode == "surge":
            records = [r for r in records if r.get("recentSurge")]
        elif ai_mode == "deplete_surge":
            records = [
                r for r in records
                if r.get("depleteWithin5y") and r.get("recentSurge")
            ]
        elif ai_mode == "mentioned":
            records = [r for r in records if r.get("aiMentioned")]

        if not records:
            if ai_mode == "mentioned" and not mentioned:
                self.chart3d.show_message(
                    "아직 AI 리포트에서 언급된 품목이 없습니다.\n"
                    "AI 분석이 완료되면 리포트에 등장한 관리번호가 여기에 표시됩니다."
                )
            else:
                self.chart3d.show_message(
                    "선택한 필터에 해당하는 3D 품목이 없습니다.\n"
                    "표준품구분/AI 연동 필터를 바꿔 보세요."
                )
            return
        self.chart3d.plot_records(
            records,
            source_file=str(self.inventory_data.get("file_name") or ""),
        )

    def _scroll_chat_to_bottom(self) -> None:
        def _go() -> None:
            sb = self.chat_view.verticalScrollBar()
            sb.setValue(sb.maximum())
        QTimer.singleShot(0, _go)

    def _show_report_find_bar(self) -> None:
        self.report_find_bar.show()
        self.report_find_input.setFocus()
        self.report_find_input.selectAll()

    def _hide_report_find_bar(self) -> None:
        self.report_find_bar.hide()
        self.report_fixed.setExtraSelections([])
        self.report_fixed.setFocus()

    def _render_report_html(self) -> None:
        """표준 리포트를 QTextBrowser용 HTML로 다시 그린다 (접기/펼치기 상태 반영)."""
        if not self._initial_report:
            self.report_fixed.clear()
            return
        self.report_fixed.setHtml(
            markdown_report_to_collapsible_html(
                self._initial_report,
                expanded_ids=self._report_expanded_ids,
            )
        )

    def _on_report_anchor_clicked(self, url: QUrl) -> None:
        """펼쳐보기/접기 앵커 처리. QTextBrowser는 <details>를 지원하지 않음."""
        frag = (url.fragment() or "").strip()
        if not frag and url.scheme() in ("expand", "collapse"):
            frag = f"{url.scheme()}:{url.path()}"
        if frag.startswith("expand:"):
            self._report_expanded_ids.add(frag.split(":", 1)[1])
            self._render_report_html()
        elif frag.startswith("collapse:"):
            self._report_expanded_ids.discard(frag.split(":", 1)[1])
            self._render_report_html()

    def _find_in_report(self, forward: bool = True) -> None:
        query = self.report_find_input.text()
        if not query:
            return
        flags = QTextDocument.FindFlag(0)
        if not forward:
            flags |= QTextDocument.FindFlag.FindBackward
        found = self.report_fixed.find(query, flags)
        if not found:
            cursor = self.report_fixed.textCursor()
            if forward:
                cursor.movePosition(QTextCursor.MoveOperation.Start)
            else:
                cursor.movePosition(QTextCursor.MoveOperation.End)
            self.report_fixed.setTextCursor(cursor)
            found = self.report_fixed.find(query, flags)
        if found:
            # 현재 선택 강조
            sel = QTextEdit.ExtraSelection()
            sel.cursor = self.report_fixed.textCursor()
            fmt = QTextCharFormat()
            fmt.setBackground(QColor("#fde68a"))
            sel.format = fmt
            self.report_fixed.setExtraSelections([sel])
        else:
            self.report_fixed.setExtraSelections([])

    def _clear_chat_history(self) -> None:
        """대화만 초기화 (좌측 표준 리포트·_initial_report 유지)."""
        self._chat_history.clear()
        if self._initial_report:
            self.chat_view.setMarkdown(
                "*표준 분석 리포트가 왼쪽에 준비되었습니다. 추가 질문을 입력해 주세요.*"
            )
        else:
            self.chat_view.clear()
        self._scroll_chat_to_bottom()

    def _set_chat_enabled(self, enabled: bool) -> None:
        self.btn_chat_send.setEnabled(enabled and not self._chat_busy)
        self.chat_input.setEnabled(enabled and not self._chat_busy)
        if hasattr(self, "btn_chat_clear"):
            self.btn_chat_clear.setEnabled(bool(self._initial_report) and not self._chat_busy)

    def _render_chat_view(self) -> None:
        parts: list[str] = []
        for msg in self._chat_history:
            role = msg.get("role", "assistant")
            text = msg.get("text", "")
            if role == "user":
                parts.append(f"### 🙋 질문\n\n{text}")
            elif role == "system":
                parts.append(f"*{text}*")
            else:
                parts.append(f"### 🤖 AI\n\n{text}")
        self.chat_view.setMarkdown("\n\n---\n\n".join(parts) if parts else "")
        self._scroll_chat_to_bottom()

    def _run_ai_analysis(self, data: dict[str, Any]) -> None:
        key = self.api_key_input.text().strip()
        stock_items = data.get("stock_items") or []
        changed = [it for it in stock_items if it.has_stock_change]
        if "ai_flags" not in data:
            data["ai_flags"] = collect_ai_analysis_flags(stock_items)
        self._chat_history.clear()
        self._initial_report = ""
        self._report_expanded_ids.clear()
        self.report_fixed.clear()
        self._set_chat_enabled(False)
        if not changed:
            self.chat_view.setMarkdown(
                "분석 대상 없음\n\n수량 변화가 있는 품목이 없어 AI 분석을 건너뛰었습니다."
            )
            self._scroll_chat_to_bottom()
            return
        if not key:
            self.chat_view.setMarkdown(
                "API Key가 설정되지 않았습니다.\n"
                "상단에서 Gemini API Key를 입력하고 [저장]을 눌러 주세요.\n\n"
                        f"(대상 품목 중 변동 {len(changed)}건 대기 중)"
            )
            self._scroll_chat_to_bottom()
            return
        flags = data.get("ai_flags") or {}
        deplete_n = len(flags.get("deplete_codes") or [])
        surge_n = len(flags.get("surge_codes") or [])
        mfg = flags.get("manufacture_candidates") or {}
        mfg_n = len(mfg.get("표준생약") or []) + len(mfg.get("지표성분") or [])
        self._chat_busy = True
        self._set_chat_enabled(False)
        self._chat_history.append(
            {
                "role": "system",
                "text": (
                    f"AI 분석 리포트 생성 중... (변동 {len(changed)}건 · "
                    f"소진후보 {deplete_n} · 가속 {surge_n} · 제조검토후보 {mfg_n})\n"
                    "최상단 KPI 대시보드·환산액·가속도 지표를 포함합니다."
                ),
            }
        )
        self._render_chat_view()
        self._show_busy("AI 분석", "초기 분석 리포트를 생성하는 중...\n잠시만 기다려 주세요.")

        def _start() -> None:
            prompt = build_ai_prompt(
                stock_items,
                compendium_context=self._compendium_prompt_context(stock_items),
                compendium_match_report=data.get("compendium_match_report") or None,
                flags=data.get("ai_flags"),
            )
            self.gemini_worker = GeminiWorker(key, prompt, followup=False)
            self.gemini_worker.finished.connect(self._on_report_ready)
            self.gemini_worker.error.connect(self._on_report_error)
            self.gemini_worker.start()

        self._run_after_busy_paint(_start)

    def _on_report_ready(self, text: str) -> None:
        self._close_busy()
        self._chat_busy = False
        self._chat_history.clear()
        self._initial_report = text
        self._report_expanded_ids.clear()
        self._render_report_html()
        self.chat_view.setMarkdown(
            "*표준 분석 리포트가 왼쪽에 준비되었습니다. 추가 질문을 입력해 주세요.*"
        )
        self._scroll_chat_to_bottom()
        self._set_chat_enabled(True)
        self._set_api_status(True, f"연결됨 ({get_active_gemini_model()})")
        if self.inventory_data is not None:
            codes = [
                it.manage_no
                for it in (self.inventory_data.get("stock_items") or [])
                if getattr(it, "manage_no", None)
            ]
            mentioned = extract_mentioned_codes_from_report(text, codes)
            self.inventory_data["ai_mentioned_codes"] = mentioned
            self._refresh_3d()

    def _on_report_error(self, message: str) -> None:
        self._close_busy()
        self._chat_busy = False
        self._chat_history = [m for m in self._chat_history if m.get("role") != "system"]
        self._chat_history.append(
            {"role": "assistant", "text": f"AI 분석 생성 실패:\n\n{message}"}
        )
        self._render_chat_view()
        self._set_chat_enabled(bool(self._initial_report))
        self._set_api_status(False, message)

    def _send_chat_message(self) -> None:
        if self._chat_busy:
            return
        question = self.chat_input.text().strip()
        if not question:
            return
        if not self._initial_report:
            QMessageBox.information(self, "안내", "초기 AI 리포트가 생성된 뒤 질문할 수 있습니다.")
            return
        key = self.api_key_input.text().strip()
        if not key:
            QMessageBox.warning(self, "경고", "API Key를 입력해 주세요.")
            return

        self.chat_input.clear()
        self._chat_history.append({"role": "user", "text": question})
        self._chat_history.append(
            {"role": "system", "text": "재고·공정서 데이터를 실시간 재검토하여 답변 생성 중..."}
        )
        self._chat_busy = True
        self._set_chat_enabled(False)
        self._render_chat_view()
        self._show_busy("AI 답변", "재고 통합 데이터·공정서 DB를 재조회하는 중...")

        def _start() -> None:
            stock_items = None
            table_df = None
            ai_flags = None
            if self.inventory_data:
                stock_items = self.inventory_data.get("stock_items")
                table_df = self.inventory_data.get("table_df")
                ai_flags = self.inventory_data.get("ai_flags")
            prompt = build_followup_prompt(
                self._initial_report,
                question,
                stock_items,
                compendium_context=self._compendium_prompt_context(stock_items),
                table_df=table_df,
                flags=ai_flags,
            )
            # 최근 대화 맥락을 짧게 첨부 (초기 리포트는 chat_history에 없음)
            prior_qas = [
                m for m in self._chat_history
                if m.get("role") in ("user", "assistant")
            ]
            if len(prior_qas) > 2:
                tail = prior_qas[-6:-1]  # 직전 질문 제외한 최근 맥락
                if tail:
                    ctx = "\n\n".join(
                        ("사용자: " if m["role"] == "user" else "AI: ") + m["text"][:800]
                        for m in tail
                    )
                    prompt = prompt + "\n\n[최근 대화 맥락]\n" + ctx

            self.gemini_worker = GeminiWorker(key, prompt, followup=True)
            self.gemini_worker.finished.connect(self._on_chat_reply)
            self.gemini_worker.error.connect(self._on_chat_error)
            self.gemini_worker.start()

        self._run_after_busy_paint(_start)

    def _on_chat_reply(self, text: str) -> None:
        self._close_busy()
        self._chat_busy = False
        self._chat_history = [m for m in self._chat_history if m.get("role") != "system"]
        self._chat_history.append({"role": "assistant", "text": text})
        self._render_chat_view()
        self._set_chat_enabled(True)
        self._set_api_status(True, f"연결됨 ({get_active_gemini_model()})")
        if self.inventory_data is not None:
            codes = [
                it.manage_no
                for it in (self.inventory_data.get("stock_items") or [])
                if getattr(it, "manage_no", None)
            ]
            # 후속 답변에서 언급된 코드도 누적
            prev = set(self.inventory_data.get("ai_mentioned_codes") or [])
            prev.update(extract_mentioned_codes_from_report(text, codes))
            self.inventory_data["ai_mentioned_codes"] = list(prev)
            self._refresh_3d()

    def _on_chat_error(self, message: str) -> None:
        self._close_busy()
        self._chat_busy = False
        self._chat_history = [m for m in self._chat_history if m.get("role") != "system"]
        self._chat_history.append(
            {"role": "assistant", "text": f"답변 생성 실패:\n\n{message}"}
        )
        self._render_chat_view()
        self._set_chat_enabled(True)
        self._set_api_status(False, message)


def create_splash(app: QApplication) -> tuple[QWidget, QProgressBar, QLabel]:
    """시작 로딩창.

    QSplashScreen + 반투명 배경은 Windows에서 안 보이거나 즉시 사라지는 경우가 많아
    불투명 프레임리스 QWidget을 사용한다.
    """
    w, h = 540, 320
    pix = QPixmap(w, h)
    pix.fill(QColor("#e8eef6"))

    painter = QPainter(pix)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    painter.setRenderHint(QPainter.RenderHint.TextAntialiasing)

    panel = QPainterPath()
    panel.addRoundedRect(14, 14, w - 28, h - 28, 18, 18)

    painter.setPen(Qt.PenStyle.NoPen)
    painter.setBrush(QBrush(QColor("#f4f7fb")))
    painter.drawPath(panel)

    highlight = QColor("#ffffff")
    highlight.setAlpha(120)
    painter.setBrush(QBrush(highlight))
    painter.drawRoundedRect(18, 18, w - 36, (h - 36) // 2, 16, 16)

    outer = QPen(QColor("#1e3a5f"))
    outer.setWidth(2)
    outer.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
    painter.setPen(outer)
    painter.setBrush(QBrush(Qt.BrushStyle.NoBrush))
    painter.drawPath(panel)

    inner = QPen(QColor("#6b87a8"))
    inner.setWidth(1)
    painter.setPen(inner)
    painter.drawRoundedRect(26, 26, w - 52, h - 52, 14, 14)

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

    painter.setPen(QColor("#0b1f3a"))
    painter.setFont(QFont("Malgun Gothic", 16, QFont.Weight.Bold))
    painter.drawText(40, 130, w - 80, 36, Qt.AlignmentFlag.AlignHCenter, "생약표준품 재고 분석 시스템")
    painter.setFont(QFont("Malgun Gothic", 11))
    painter.setPen(QColor("#3d5a80"))
    painter.drawText(40, 168, w - 80, 28, Qt.AlignmentFlag.AlignHCenter, f"{APP_VERSION} 초기화 중...")
    painter.end()

    splash = QWidget(None, Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
    splash.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose, True)
    splash.setFixedSize(w, h)
    splash.setWindowTitle(f"생약표준품 재고 분석 시스템 {APP_VERSION}")
    splash.setStyleSheet("background-color: #e8eef6;")

    bg = QLabel(splash)
    bg.setPixmap(pix)
    bg.setGeometry(0, 0, w, h)
    bg.lower()

    status = QLabel(splash)
    status.setGeometry(40, 214, w - 80, 24)
    status.setAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
    status.setStyleSheet(
        "color: #3d5a80; font-size: 12px; font-family: 'Malgun Gothic'; background: transparent;"
    )
    status.setText("모듈 로드 중...")

    bar = QProgressBar(splash)
    bar.setGeometry(70, 246, w - 140, 16)
    bar.setRange(0, 100)
    bar.setValue(8)
    bar.setTextVisible(False)
    bar.setStyleSheet(
        """
        QProgressBar {
            background: #ffffff;
            border: 1px solid #1e3a5f;
            border-radius: 8px;
        }
        QProgressBar::chunk {
            background-color: #1e3a5f;
            border-radius: 7px;
            margin: 2px;
        }
        """
    )

    # 화면 중앙 배치
    screen = app.primaryScreen()
    if screen is not None:
        geo = screen.availableGeometry()
        splash.move(geo.center().x() - w // 2, geo.center().y() - h // 2)

    splash.show()
    splash.raise_()
    splash.activateWindow()
    app.processEvents()
    return splash, bar, status


def main() -> None:
    # QWebEngineView 사용 시 QApplication 생성 전에 필요
    try:
        QApplication.setAttribute(Qt.ApplicationAttribute.AA_ShareOpenGLContexts)
    except Exception:
        pass
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(APP_STYLE)

    splash, bar, status_label = create_splash(app)
    app.processEvents()

    # 스플래시가 먼저 보이도록 메인 윈도우 생성은 약간 지연
    state = {"i": 0, "window": None}
    steps = [15, 35, 55, 75, 90, 100]
    labels = [
        "모듈 로드 중...",
        "UI 구성 중...",
        "차트 엔진 준비 중...",
        "AI 연동 준비 중...",
        "마무리 중...",
        f"생약표준품 재고 분석 시스템 {APP_VERSION} 초기화 중...",
    ]

    def tick() -> None:
        i = state["i"]
        if i == 1 and state["window"] is None:
            status_label.setText(labels[1])
            bar.setValue(steps[1])
            splash.raise_()
            app.processEvents()
            state["window"] = MainWindow()
            state["i"] = 2
            QTimer.singleShot(280, tick)
            return
        if i < len(steps):
            bar.setValue(steps[i])
            status_label.setText(labels[i])
            splash.raise_()
            app.processEvents()
            state["i"] += 1
            QTimer.singleShot(320, tick)
        else:
            window = state["window"] or MainWindow()
            window.show()
            window.raise_()
            window.activateWindow()
            splash.close()

    QTimer.singleShot(120, tick)
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
