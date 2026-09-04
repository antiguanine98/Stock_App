"""
생약표준품 재고 분석 및 소급 보정 시스템 (PyQt6) v1.65
"""

from __future__ import annotations

import json
import os
import sys
import traceback
from pathlib import Path
from typing import Any, Callable
from urllib.parse import unquote

import matplotlib
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.ticker import MaxNLocator
from PyQt6.QtCore import QEvent, QMarginsF, QObject, Qt, QThread, QTimer, QUrl, pyqtSignal
from PyQt6.QtGui import (
    QBrush,
    QColor,
    QDragEnterEvent,
    QDropEvent,
    QFont,
    QKeySequence,
    QPageLayout,
    QPageSize,
    QPainter,
    QPainterPath,
    QPdfWriter,
    QPen,
    QPixmap,
    QShortcut,
    QTextCharFormat,
    QTextCursor,
    QTextDocument,
)
from PyQt6.QtWidgets import (
    QApplication,
    QButtonGroup,
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
    attach_compendium_match_to_flags,
    build_ai_prompt,
    build_followup_prompt,
    build_scatter3d_records,
    collect_ai_analysis_flags,
    ensure_mandatory_report_sections,
    export_markdown_report_to_docx,
    extract_mentioned_codes_from_report,
    format_compendium_context,
    format_compendium_match_report,
    format_qty_int,
    load_compendium_excel,
    lookup_pharmacopoeia_tag,
    markdown_report_to_collapsible_html,
    match_compendium_inventory,
    normalize_excel_path,
    process_excel,
    process_excels,
    split_markdown_report_sections,
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
APP_VERSION = "v1.65"
AUTHOR_CREDIT = "made by 2026MFDSyouthinternKYHLCY"

# 서버 확인 최신 Flash — 탐색 실패 시에도 이 기본값으로 연결
GEMINI_DEFAULT_MODEL = "gemini-3.6-flash"
GEMINI_MODEL_PREFERENCES = [
    "gemini-3.6-flash",
    "gemini-2.5-flash",
]
GEMINI_FAILOVER_PREFERENCES = [
    "gemini-3.6-flash",
    "gemini-2.5-flash",
]
# 할당량 0·종료·미지원 모델 (omni 등)
GEMINI_RETIRED_MODELS = frozenset(
    {
        "gemini-omni-1.1-flash",
        "gemini-omni-flash",
        "gemini-1.5-flash",
        "gemini-1.5-flash-001",
        "gemini-2.0-flash",
        "gemini-2.0-flash-lite",
        "gemini-2.0-flash-001",
        "gemini-2.0-flash-lite-001",
        "gemini-2.5-flash-lite",
        "gemini-2.5-flash-lite-preview",
        "gemini-2.5-flash-preview",
    }
)
_resolved_gemini_model: str | None = None
_MAX_MODEL_PROBES = 6
_MAX_CASCADE_MODELS = 1  # 분석은 sticky/기본 1개 + 동일 모델 재시도
_cancel_gemini: bool = False
_discovered_flash_cache: list[str] | None = None

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
QLabel#stepLabel {
    color: #0b1f3a;
    font-size: 13px;
    font-weight: 700;
    padding: 2px 0;
}
QLabel#apiBadge {
    font-size: 13px;
    font-weight: 700;
    padding: 4px 10px;
    border-radius: 8px;
    background: #f1f5f9;
    min-width: 88px;
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
QPushButton#primaryBtn {
    background-color: #0f766e;
    color: #ffffff;
    border: none;
    border-radius: 10px;
    font-weight: 700;
    font-size: 14px;
    padding: 10px 22px;
}
QPushButton#primaryBtn:hover { background-color: #0d9488; }
QPushButton#primaryBtn:disabled {
    background-color: #94a3b8;
    color: #e2e8f0;
}
QPushButton#reportNavBtn {
    background-color: #eef3f9;
    color: #1e3a5f;
    border: 1px solid #c5d0de;
    border-radius: 10px;
    min-width: 58px;
    max-width: 64px;
    min-height: 40px;
    padding: 8px 4px;
    font-weight: 700;
    font-size: 12px;
    text-align: center;
}
QPushButton#reportNavBtn:hover {
    background-color: #dce6f2;
    border-color: #1e3a5f;
}
QPushButton#reportNavBtn:checked {
    background-color: #1e3a5f;
    color: #ffffff;
    border-color: #152a45;
}
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
    from google.genai import types

    # SDK 자체 재시도와 앱 cascade가 중첩되면 수분간 "과부하"에 고착됨 → SDK 재시도 최소화
    try:
        http_opts = types.HttpOptions(
            timeout=_HTTP_TIMEOUT_MS,
            retry_options=types.HttpRetryOptions(
                attempts=1,
                initial_delay=0.5,
                max_delay=2.0,
            ),
        )
    except Exception:
        try:
            http_opts = types.HttpOptions(timeout=_HTTP_TIMEOUT_MS)
        except Exception:
            http_opts = None
    if http_opts is not None:
        return genai.Client(api_key=api_key, http_options=http_opts)
    return genai.Client(api_key=api_key)


def _api_error_code(exc: BaseException) -> int | None:
    try:
        from google.genai import errors as genai_errors

        if isinstance(exc, genai_errors.APIError):
            code = getattr(exc, "code", None)
            if code is not None:
                return int(code)
    except (ImportError, TypeError, ValueError):
        pass
    # 문자열에서 HTTP 코드 추출 (래핑·pickle 복원 케이스)
    text = str(exc)
    import re as _re

    m = _re.search(r"\b(429|500|502|503|504|401|403|404)\b", text)
    if m:
        return int(m.group(1))
    low = text.lower()
    if "unavailable" in low or "overloaded" in low:
        return 503
    if "resource_exhausted" in low or "rate limit" in low:
        return 429
    return None


def _is_auth_error(exc: BaseException) -> bool:
    code = _api_error_code(exc)
    if code in (401, 403):
        return True
    text = str(exc).lower()
    return any(k in text for k in ("api key", "api_key", "permission denied", "unauthenticated"))


_RETRYABLE_CODES = (429, 500, 502, 503, 504)
_MAX_RETRIES = 2  # 동일 모델 최대 2회 재시도
_RETRY_BASE_SECONDS = 3  # 실패 시 고정 3초 대기
_RETRY_MAX_WAIT = 3
_HTTP_TIMEOUT_MS = 120_000  # 분석 요청 타임아웃(ms) — ReadTimeout 완화


def request_cancel_gemini() -> None:
    """백그라운드 Gemini 호출 취소를 요청한다."""
    global _cancel_gemini
    _cancel_gemini = True


def clear_cancel_gemini() -> None:
    global _cancel_gemini
    _cancel_gemini = False


def _sleep_with_cancel(seconds: float) -> None:
    """취소 가능한 대기 (0.25초 단위 폴링)."""
    import time

    end = time.monotonic() + max(0.0, seconds)
    while time.monotonic() < end:
        if _cancel_gemini:
            raise RuntimeError("사용자에 의해 AI 요청이 취소되었습니다.")
        time.sleep(min(0.25, end - time.monotonic()))


def _is_retryable_error(exc: BaseException) -> bool:
    code = _api_error_code(exc)
    if code in _RETRYABLE_CODES:
        return True
    text = str(exc).lower()
    return any(
        m in text
        for m in (
            "resource_exhausted",
            "rate limit",
            "quota",
            "unavailable",
            "high demand",
            "overloaded",
            "overload",
            "deadline",
            "timed out",
            "timeout",
            "temporarily",
            "try again",
            "server error",
            "internal error",
        )
    )


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


def _is_retired_model(model_id: str | None) -> bool:
    """할당량 0·종료·미지원 모델인지 (omni 등)."""
    if not model_id:
        return False
    mid = str(model_id).strip()
    if mid in GEMINI_RETIRED_MODELS:
        return True
    low = mid.lower()
    if "omni" in low:
        return True
    if low.startswith("gemini-2.0-flash"):
        return True
    if low.startswith("gemini-1.5-flash") and low != "gemini-1.5-flash-latest":
        return True
    return False


def _normalize_model_id(name: str | None) -> str:
    """models/gemini-… → gemini-…"""
    mid = str(name or "").strip()
    if "/" in mid:
        mid = mid.split("/")[-1].strip()
    return mid


def _model_supports_generate_content(model_obj: Any) -> bool:
    """generateContent 지원 여부 — 필드 비어 있으면 True(관대)."""
    actions = (
        getattr(model_obj, "supported_actions", None)
        or getattr(model_obj, "supported_generation_methods", None)
        or []
    )
    if not actions:
        return True
    for raw in actions:
        a = str(raw).lower().replace("_", "")
        if "generatecontent" in a:
            return True
    return False


def _list_generate_content_models(client) -> list[str]:
    """API 목록에서 generateContent 가능 모델 ID (대소문자 무시·관대 필터)."""
    available: list[str] = []
    try:
        pager = client.models.list()
        for m in pager:
            name = str(getattr(m, "name", "") or "")
            mid = _normalize_model_id(name)
            if not mid:
                continue
            if not _model_supports_generate_content(m):
                continue
            low = mid.lower()
            # 임베딩·음성·실시간·이미지·omni(할당량0)만 제외
            if any(x in low for x in ("embed", "tts", "live", "imagen", "image", "omni")):
                continue
            if _is_retired_model(mid):
                continue
            if mid not in available:
                available.append(mid)
    except Exception as exc:
        log_gemini(
            "WARN",
            f"models.list 실패(기본 모델로 진행): {format_gemini_error(exc)}",
        )
    return available


def _pick_target_from_available(available: list[str]) -> str | None:
    """1) 3.6-flash 2) flash 포함 3) 목록 첫 항목."""
    for m_name in available:
        if "3.6-flash" in m_name.lower():
            return m_name
    for m_name in available:
        if "flash" in m_name.lower():
            return m_name
    if available:
        return available[0]
    return None


def _discover_flash_models(client, *, use_cache: bool = True) -> list[str]:
    """연결/분석용 후보 — 3.6-flash 우선, 실패 시 기본값 보장."""
    global _discovered_flash_cache
    if use_cache and _discovered_flash_cache is not None:
        return list(_discovered_flash_cache)

    available = _list_generate_content_models(client)
    ordered: list[str] = []

    def _add(mid: str | None) -> None:
        if mid and not _is_retired_model(mid) and mid not in ordered:
            ordered.append(mid)

    # 우선순위 매핑
    target = _pick_target_from_available(available)
    _add(target)
    for preferred in GEMINI_MODEL_PREFERENCES:
        _add(preferred)
    for mid in available:
        if "flash" in mid.lower():
            _add(mid)
    for mid in available:
        _add(mid)
    # 리스트 조회 실패·비어도 임의 에러 없이 안정 기본값
    if not ordered:
        _add(GEMINI_DEFAULT_MODEL)
    _add(GEMINI_DEFAULT_MODEL)

    _discovered_flash_cache = list(ordered)
    log_gemini(
        "INFO",
        f"모델 후보 {len(ordered)}개 (1순위={ordered[0]}): "
        + ", ".join(ordered[:8])
        + ("…" if len(ordered) > 8 else ""),
    )
    return ordered


def _pick_latest_flash(discovered: list[str] | None = None) -> str | None:
    disc = [m for m in (discovered or []) if not _is_retired_model(m)]
    if disc:
        picked = _pick_target_from_available(disc)
        if picked:
            return picked
    return GEMINI_DEFAULT_MODEL


def _candidate_models(prioritize: str | None = None, *, discovered: list[str] | None = None) -> list[str]:
    """연결 테스트용 후보 — 탐색 결과 + 기본값 (비어 있지 않음)."""
    ordered: list[str] = []

    def _add(mid: str | None) -> None:
        if mid and not _is_retired_model(mid) and mid not in ordered:
            ordered.append(mid)

    _add(prioritize)
    _add(_resolved_gemini_model)
    disc = list(discovered or [])
    target = _pick_target_from_available(disc)
    _add(target)
    for preferred in GEMINI_MODEL_PREFERENCES:
        _add(preferred)
    for mid in disc:
        _add(mid)
    _add(GEMINI_DEFAULT_MODEL)
    return ordered[:_MAX_MODEL_PROBES] or [GEMINI_DEFAULT_MODEL]


def _cascade_models(primary: str | None, discovered: list[str] | None = None) -> list[str]:
    """분석용 — sticky/기본 1개 (omni 페일오버 금지)."""
    ordered: list[str] = []

    def _add(mid: str | None) -> None:
        if mid and not _is_retired_model(mid) and mid not in ordered:
            ordered.append(mid)

    _add(primary)
    _add(_pick_target_from_available(list(discovered or [])))
    _add(GEMINI_DEFAULT_MODEL)
    for mid in GEMINI_MODEL_PREFERENCES:
        _add(mid)
        if len(ordered) >= _MAX_CASCADE_MODELS:
            break
    return ordered[:_MAX_CASCADE_MODELS] or [GEMINI_DEFAULT_MODEL]


def extract_response_text(response: Any) -> str | None:
    text = getattr(response, "text", None)
    if text and str(text).strip():
        return str(text).strip()
    return None


def _probe_model(client, model: str) -> str:
    """가벼운 핑 테스트.

    Returns:
        "ok"         — 정상 응답
        "overloaded" — 키/모델은 유효하나 503/429 등 일시 과부하
        "fail"       — 모델 없음·기타 실패
    """
    from google.genai import types

    try:
        try:
            config = types.GenerateContentConfig(
                max_output_tokens=32,
                temperature=0,
                thinking_config=types.ThinkingConfig(thinking_budget=0),
            )
        except Exception:
            config = types.GenerateContentConfig(max_output_tokens=32, temperature=0)
        response = client.models.generate_content(
            model=model,
            contents="ping",
            config=config,
        )
        if extract_response_text(response) is not None:
            return "ok"
        candidates = getattr(response, "candidates", None) or []
        return "ok" if candidates else "fail"
    except Exception as exc:
        if _is_auth_error(exc):
            raise
        detail = format_gemini_error(exc)
        if _is_retryable_error(exc):
            log_gemini("WARN", f"모델 일시 과부하: {model} ({detail})")
            return "overloaded"
        log_gemini("WARN", f"모델 핑 실패, 다음 후보: {model} ({detail})")
        return "fail"


def resolve_gemini_model(client, force_refresh: bool = False) -> str:
    """연결에 쓸 모델 선택 — 탐색 실패해도 gemini-3.6-flash 기본값으로 연결.

    핑이 ok/overloaded면 채택. 후보 전부가 fail이어도 기본 모델을 sticky로
    두고 연결 성공 처리한다 (미연결 오판 방지).
    """
    global _resolved_gemini_model
    if _resolved_gemini_model and not force_refresh and not _is_retired_model(
        _resolved_gemini_model
    ):
        return _resolved_gemini_model

    discovered = _discover_flash_models(client, use_cache=not force_refresh)
    candidates = _candidate_models(
        prioritize=_resolved_gemini_model, discovered=discovered
    )
    log_gemini("INFO", f"연결 후보: {', '.join(candidates)}")

    errors: list[str] = []
    overloaded: list[str] = []
    for model in candidates:
        log_gemini("INFO", f"모델 핑: {model}")
        try:
            status = _probe_model(client, model)
            if status == "ok":
                _resolved_gemini_model = model
                log_gemini("INFO", f"사용 모델 선택: {model}")
                return model
            if status == "overloaded":
                overloaded.append(model)
                errors.append(f"{model}: 일시 과부하")
                continue
            errors.append(f"{model}: 핑 실패")
        except Exception as exc:
            if _is_auth_error(exc):
                raise
            detail = format_gemini_error(exc)
            errors.append(f"{model}: {detail}")
            log_gemini("ERROR", f"모델 검증 실패: {model} — {detail}", exc)

    if overloaded:
        pick = overloaded[0]
        _resolved_gemini_model = pick
        log_gemini("INFO", f"과부하 상태지만 모델 채택(연결 유지): {pick}")
        return pick

    # 탐색/핑이 모두 실패해도 기본값으로 연결 (미연결 런타임 에러 방지)
    pick = GEMINI_DEFAULT_MODEL
    _resolved_gemini_model = pick
    log_gemini(
        "WARN",
        "후보 핑 실패 — 기본 모델로 연결 유지: "
        f"{pick} | "
        + ("; ".join(errors[:4]) if errors else "후보 없음"),
    )
    return pick


def get_active_gemini_model() -> str:
    if _resolved_gemini_model and not _is_retired_model(_resolved_gemini_model):
        return _resolved_gemini_model
    pick = _pick_latest_flash(_discovered_flash_cache)
    return pick or GEMINI_DEFAULT_MODEL


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
    """API Key 연결 테스트 — 성공 시 '연결됨 (모델명)'."""
    global _resolved_gemini_model, _discovered_flash_cache
    client = create_gemini_client(api_key)
    log_gemini("INFO", "연결 테스트 시작")
    _discovered_flash_cache = None
    if _is_retired_model(_resolved_gemini_model):
        _resolved_gemini_model = None
    if _resolved_gemini_model:
        status = _probe_model(client, _resolved_gemini_model)
        if status == "ok":
            return True, f"연결됨 ({_resolved_gemini_model})"
        if status == "overloaded":
            return True, f"연결됨 ({_resolved_gemini_model}) — 서버 일시 과부하"
        _resolved_gemini_model = None
    model = resolve_gemini_model(client, force_refresh=True)
    return True, f"연결됨 ({model})"


_retry_stage_callback: Callable[[str], None] | None = None


def _call_with_retry(
    client,
    model: str,
    prompt: str,
    *,
    max_retries: int | None = None,
) -> Any:
    """generate_content — 실패 시 3초 대기 후 최대 2회 재시도 (동일 모델)."""
    from google.genai import types

    retries = _MAX_RETRIES if max_retries is None else max(0, int(max_retries))
    try:
        gen_config = types.GenerateContentConfig(
            temperature=0.15,
            thinking_config=types.ThinkingConfig(thinking_budget=0),
        )
    except Exception:
        try:
            gen_config = types.GenerateContentConfig(temperature=0.15)
        except Exception:
            gen_config = None

    last_exc: BaseException | None = None
    for attempt in range(retries + 1):
        if _cancel_gemini:
            raise RuntimeError("사용자에 의해 AI 요청이 취소되었습니다.")
        try:
            if gen_config is not None:
                return client.models.generate_content(
                    model=model, contents=prompt, config=gen_config
                )
            return client.models.generate_content(model=model, contents=prompt)
        except Exception as exc:
            if _is_auth_error(exc):
                raise
            if not _is_retryable_error(exc):
                raise
            last_exc = exc
            if attempt < retries:
                wait = float(_RETRY_BASE_SECONDS)
                msg = (
                    f"일시 오류({model}) — {wait:.0f}초 후 재시도 "
                    f"({attempt + 1}/{retries})..."
                )
                log_gemini("WARN", msg)
                if _retry_stage_callback:
                    _retry_stage_callback(msg)
                _sleep_with_cancel(wait)
    raise last_exc  # type: ignore[misc]


def generate_gemini_report(api_key: str, prompt: str) -> str:
    """Gemini 분석 호출 — sticky/기본(gemini-3.6-flash), 실패 시 3초×최대 2회 재시도.

    omni 등 할당량 0 모델로의 페일오버는 하지 않는다.
    """
    global _resolved_gemini_model
    client = create_gemini_client(api_key)
    if _is_retired_model(_resolved_gemini_model):
        log_gemini("WARN", f"retired 모델 sticky 제거: {_resolved_gemini_model}")
        _resolved_gemini_model = None
    model = (
        _resolved_gemini_model
        if _resolved_gemini_model and not _is_retired_model(_resolved_gemini_model)
        else GEMINI_DEFAULT_MODEL
    )

    log_gemini("INFO", f"AI 분석 요청 (model={model}, retries≤{_MAX_RETRIES})")
    if _retry_stage_callback:
        _retry_stage_callback(f"AI 분석 중… ({model})")
    try:
        response = _call_with_retry(client, model, prompt, max_retries=_MAX_RETRIES)
        out = extract_response_text(response)
        if out:
            _resolved_gemini_model = model
            log_gemini("INFO", f"AI 응답 성공 (model={model})")
            return out
        raise RuntimeError(describe_empty_response(response))
    except Exception as exc:
        if _is_auth_error(exc):
            raise
        if "취소" in str(exc):
            raise
        detail = format_gemini_error(exc)
        low = detail.lower()
        if "quota" in low or "resource_exhausted" in low or "429" in low:
            raise RuntimeError(
                "Gemini API 할당량(토큰/요청 한도)을 초과했습니다.\n"
                "프롬프트는 요약 데이터만 전송하도록 축소되어 있습니다. "
                "잠시 후 다시 시도하거나 AI Studio 할당량을 확인해 주세요.\n"
                f"{detail}"
            ) from exc
        raise RuntimeError(
            "AI 분석 요청에 실패했습니다.\n"
            f"모델: {model} (재시도 {_MAX_RETRIES}회 포함)\n"
            "대기 중에는 [중단]으로 UI를 해제할 수 있습니다.\n\n"
            f"{detail}"
        ) from exc


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
            self.finished.emit(process_excels(self.file_paths))
        except Exception as e:
            self.error.emit(str(e))


class GeminiWorker(QThread):
    """프롬프트 구성 + Gemini 호출을 백그라운드에서 수행 (UI 논블로킹)."""

    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    stage = pyqtSignal(str)

    def __init__(
        self,
        api_key: str,
        *,
        prompt: str | None = None,
        build_prompt: Callable[[], str] | None = None,
        followup: bool = False,
    ):
        super().__init__()
        self.api_key = api_key
        self.prompt = prompt
        self.build_prompt = build_prompt
        self.followup = followup

    def run(self) -> None:
        global _retry_stage_callback
        clear_cancel_gemini()
        _retry_stage_callback = lambda msg: self.stage.emit(msg)
        try:
            if self.build_prompt is not None:
                self.stage.emit("재고 및 공정서 DB 분석 중...")
                prompt = self.build_prompt()
            else:
                prompt = self.prompt or ""
            self.stage.emit(
                "AI 답변 생성 중..." if self.followup else "초기 분석 리포트 생성 중..."
            )
            self.finished.emit(generate_gemini_report(self.api_key, prompt))
        except Exception as e:
            detail = format_gemini_error(e)
            kind = "후속 질문" if self.followup else "AI 분석"
            log_gemini("ERROR", f"{kind} 실패: {detail}", e)
            self.error.emit(detail)
        finally:
            _retry_stage_callback = None


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
            raw = url.toLocalFile() if hasattr(url, "toLocalFile") else url
            path = normalize_excel_path(raw)
            if path:
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
            f"{label}\n연도: {y}\n재고량: {format_qty_int(q)}" if q == q else f"{label}\n연도: {y}"
            for y, q in zip(years, y_corr)
        ]
        self._attach_hover(ax, line, labels)

        ax.set_xticks(years)
        ax.set_xticklabels([str(y) for y in years], rotation=0, ha="center", fontsize=9)
        ax.set_xlabel("연도 (YYYY)")
        ax.set_ylabel("재고량")
        ax.yaxis.set_major_locator(MaxNLocator(integer=True))
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
                    f"{item['label']}\n연도: {yr}\n재고량: {format_qty_int(q)}" if q == q else f"{item['label']}\n연도: {yr}"
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
                        qtxt = format_qty_int(q) if q is not None else "-"
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
        ax.yaxis.set_major_locator(MaxNLocator(integer=True))
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
        self._pending_excel_paths: list[str] = []
        self._chat_history: list[dict[str, str]] = []
        self._initial_report: str = ""
        self._report_sections: list[dict[str, str]] = []
        self._report_section_key: str = "all"  # "all" | section id
        self._report_expanded_ids: set[str] = set()
        self._chat_busy = False
        self.compendium_df = None
        self.compendium_meta: dict[str, Any] | None = None
        self.compendium_worker: CompendiumWorker | None = None
        self._busy_dialog: QProgressDialog | None = None
        self._api_connected = False

        self._build_ui()
        # API Key는 세션 메모리만 사용 — 파일/레지스트리에서 불러오지 않음
        self._update_api_badge(False, "미연결")

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
        top_layout.setSpacing(12)

        # --- STEP 1: Gemini API (최상단) ---
        step1 = QLabel("STEP 1 · Gemini API 설정")
        step1.setObjectName("stepLabel")
        top_layout.addWidget(step1)

        api_row = QHBoxLayout()
        api_row.setSpacing(8)
        self.api_key_input = QLineEdit()
        self.api_key_input.setPlaceholderText("Google Gemini API Key (이 실행 세션에서만 유지)")
        self.api_key_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.api_key_input.textChanged.connect(self._on_api_key_edited)
        api_row.addWidget(self.api_key_input, stretch=1)
        self.btn_test = QPushButton("연결 테스트")
        self.btn_test.setObjectName("secondaryBtn")
        self.btn_test.setMinimumHeight(36)
        self.btn_test.clicked.connect(self._test_api_connection)
        api_row.addWidget(self.btn_test)
        self.api_badge = QLabel("🔴 미연결")
        self.api_badge.setObjectName("apiBadge")
        self.api_badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.api_badge.setToolTip("STEP 1에서 연결 테스트가 성공해야 분석·챗봇을 사용할 수 있습니다.")
        api_row.addWidget(self.api_badge)
        top_layout.addLayout(api_row)
        api_hint = QLabel("API Key는 디스크에 저장되지 않으며, 프로그램을 종료하면 사라집니다.")
        api_hint.setStyleSheet("color: #64748b; font-size: 11px;")
        top_layout.addWidget(api_hint)

        # --- STEP 2: 데이터 파일 (재고 | 공정서 나란히) ---
        step2 = QLabel("STEP 2 · 데이터 파일 등록")
        step2.setObjectName("stepLabel")
        top_layout.addWidget(step2)

        files_row = QHBoxLayout()
        files_row.setSpacing(12)

        stock_col = QVBoxLayout()
        stock_col.setSpacing(6)
        file_label = QLabel("재고 엑셀 업로드")
        file_label.setObjectName("sectionLabel")
        stock_col.addWidget(file_label)
        stock_inner = QHBoxLayout()
        self.dropzone = DropZone(min_height=78)
        self.dropzone.files_dropped.connect(self._load_excels)
        stock_inner.addWidget(self.dropzone, stretch=1)
        btn_col = QVBoxLayout()
        btn_select = QPushButton("파일 선택")
        btn_select.setObjectName("secondaryBtn")
        btn_select.setMinimumWidth(100)
        btn_select.setMinimumHeight(42)
        btn_select.clicked.connect(self._browse_file)
        btn_clear = QPushButton("목록 초기화")
        btn_clear.setObjectName("secondaryBtn")
        btn_clear.setMinimumWidth(100)
        btn_clear.clicked.connect(self._clear_loaded_files)
        btn_col.addWidget(btn_select)
        btn_col.addWidget(btn_clear)
        stock_inner.addLayout(btn_col)
        stock_col.addLayout(stock_inner)
        files_row.addLayout(stock_col, stretch=1)

        comp_col = QVBoxLayout()
        comp_col.setSpacing(6)
        comp_label = QLabel("공정서 DB 업로드 (규격 참조)")
        comp_label.setObjectName("sectionLabel")
        comp_col.addWidget(comp_label)
        comp_inner = QHBoxLayout()
        self.compendium_dropzone = DropZone(
            title="공정서 DB 엑셀을 드래그 앤 드롭",
            subtitle=".xlsx / .xls · AI 규격/기준 참조 전용",
            min_height=78,
        )
        self.compendium_dropzone.files_dropped.connect(self._on_compendium_dropped)
        comp_inner.addWidget(self.compendium_dropzone, stretch=1)
        comp_btn_col = QVBoxLayout()
        btn_comp_select = QPushButton("공정서 선택")
        btn_comp_select.setObjectName("secondaryBtn")
        btn_comp_select.setMinimumWidth(100)
        btn_comp_select.clicked.connect(self._browse_compendium)
        btn_comp_clear = QPushButton("공정서 해제")
        btn_comp_clear.setObjectName("secondaryBtn")
        btn_comp_clear.setMinimumWidth(100)
        btn_comp_clear.clicked.connect(self._clear_compendium)
        comp_btn_col.addWidget(btn_comp_select)
        comp_btn_col.addWidget(btn_comp_clear)
        comp_inner.addLayout(comp_btn_col)
        comp_col.addLayout(comp_inner)
        files_row.addLayout(comp_col, stretch=1)

        top_layout.addLayout(files_row)
        self.compendium_status = QLabel("공정서 DB: 미등록")
        self.compendium_status.setStyleSheet("color: #64748b; font-size: 12px;")
        top_layout.addWidget(self.compendium_status)

        analyze_row = QHBoxLayout()
        analyze_row.addStretch(1)
        self.btn_analyze = QPushButton("🚀 분석 시작")
        self.btn_analyze.setObjectName("primaryBtn")
        self.btn_analyze.setMinimumHeight(44)
        self.btn_analyze.setMinimumWidth(180)
        self.btn_analyze.setEnabled(False)
        self.btn_analyze.setToolTip(
            "재고 엑셀과 공정서 DB를 모두 등록한 뒤 클릭하세요. "
            "분석 시작 시 상단 업로드 영역이 접힙니다."
        )
        self.btn_analyze.clicked.connect(self._on_analyze_clicked)
        analyze_row.addWidget(self.btn_analyze)
        analyze_row.addStretch(1)
        top_layout.addLayout(analyze_row)
        self.analyze_hint = QLabel(
            "재고 엑셀 + 공정서 DB 등록 후 [🚀 분석 시작]을 누르면 AI 분석이 실행됩니다."
        )
        self.analyze_hint.setStyleSheet("color: #64748b; font-size: 11px;")
        self.analyze_hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        top_layout.addWidget(self.analyze_hint)

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
            "행을 더블클릭하면 해당 품목의 분양차트로 이동합니다. "
            "(등록일자 제외 · 잔고→재고 · 변경일자는 연도 YYYY · 재고량은 정수)"
        )
        hint.setStyleSheet("color: #64748b; font-size: 12px;")
        table_layout.addWidget(hint)
        self._table_find_matches: list[tuple[int, int]] = []
        self._table_find_index = -1
        self.table_find_bar = QWidget()
        table_find_row = QHBoxLayout(self.table_find_bar)
        table_find_row.setContentsMargins(0, 0, 0, 4)
        table_find_row.setSpacing(6)
        self.table_find_input = QLineEdit()
        self.table_find_input.setPlaceholderText("관리번호·생약명·성분명 검색…")
        self.table_find_input.returnPressed.connect(lambda: self._find_in_table(forward=True))
        table_find_row.addWidget(self.table_find_input, stretch=1)
        self.btn_table_find_prev = QPushButton("이전")
        self.btn_table_find_prev.clicked.connect(lambda: self._find_in_table(forward=False))
        table_find_row.addWidget(self.btn_table_find_prev)
        self.btn_table_find_next = QPushButton("다음")
        self.btn_table_find_next.clicked.connect(lambda: self._find_in_table(forward=True))
        table_find_row.addWidget(self.btn_table_find_next)
        self.btn_table_find_close = QPushButton("닫기")
        self.btn_table_find_close.clicked.connect(self._hide_table_find_bar)
        table_find_row.addWidget(self.btn_table_find_close)
        self.table_find_bar.hide()
        table_layout.addWidget(self.table_find_bar)
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
        self._table_find_shortcut = QShortcut(QKeySequence.StandardKey.Find, self)
        self._table_find_shortcut.activated.connect(self._on_global_find_shortcut)
        self._table_find_esc = QShortcut(QKeySequence(Qt.Key.Key_Escape), self.table_find_bar)
        self._table_find_esc.activated.connect(self._hide_table_find_bar)
        self.tabs.addTab(table_wrap, "분양현황")

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
        self.tabs.addTab(chart_widget, "분양차트")

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
        title_row = QWidget()
        title_row_layout = QHBoxLayout(title_row)
        title_row_layout.setContentsMargins(0, 0, 0, 0)
        title_row_layout.setSpacing(8)
        left_title = QLabel("표준 분석 리포트")
        left_title.setObjectName("sectionLabel")
        title_row_layout.addWidget(left_title, stretch=1)
        self.btn_dl_word = QPushButton("📄 Word 다운로드")
        self.btn_dl_word.setToolTip("AI 리포트를 Word(.docx)로 저장")
        self.btn_dl_word.clicked.connect(self._download_report_word)
        self.btn_dl_word.setEnabled(False)
        title_row_layout.addWidget(self.btn_dl_word)
        self.btn_dl_pdf = QPushButton("📑 PDF 다운로드")
        self.btn_dl_pdf.setToolTip("현재 리포트 서식을 PDF로 저장")
        self.btn_dl_pdf.clicked.connect(self._download_report_pdf)
        self.btn_dl_pdf.setEnabled(False)
        title_row_layout.addWidget(self.btn_dl_pdf)
        left_layout.addWidget(title_row)

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

        report_body = QWidget()
        report_body_row = QHBoxLayout(report_body)
        report_body_row.setContentsMargins(0, 0, 0, 0)
        report_body_row.setSpacing(8)

        # 좌측 섹션 버튼 (고정 라벨: 전체·요약·소진·미보유·검토…)
        nav_wrap = QWidget()
        nav_wrap.setFixedWidth(72)
        nav_outer = QVBoxLayout(nav_wrap)
        nav_outer.setContentsMargins(0, 0, 0, 0)
        nav_outer.setSpacing(0)
        self.report_nav_scroll = QScrollArea()
        self.report_nav_scroll.setWidgetResizable(True)
        self.report_nav_scroll.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        self.report_nav_scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.report_nav_scroll.setStyleSheet(
            "QScrollArea { background: transparent; border: none; }"
        )
        self.report_nav_host = QWidget()
        self.report_nav_layout = QVBoxLayout(self.report_nav_host)
        self.report_nav_layout.setContentsMargins(2, 2, 2, 2)
        self.report_nav_layout.setSpacing(6)
        self.report_nav_layout.addStretch(1)
        self.report_nav_scroll.setWidget(self.report_nav_host)
        nav_outer.addWidget(self.report_nav_scroll, stretch=1)
        report_body_row.addWidget(nav_wrap, stretch=0)

        self._report_nav_group = QButtonGroup(self)
        self._report_nav_group.setExclusive(True)
        self._report_nav_group.idClicked.connect(self._on_report_nav_clicked)
        self._report_nav_buttons: dict[str, QPushButton] = {}

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
            "QTextBrowser { padding: 12px; line-height: 1.65; font-family: 'Malgun Gothic'; font-size: 10.5pt; "
            "border: 1px solid #d8e0ea; border-radius: 10px; background: #ffffff; }"
        )
        report_body_row.addWidget(self.report_fixed, stretch=1)
        left_layout.addWidget(report_body, stretch=1)

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

        self.chat_busy_bar = QWidget()
        chat_busy_layout = QHBoxLayout(self.chat_busy_bar)
        chat_busy_layout.setContentsMargins(2, 0, 2, 0)
        chat_busy_layout.setSpacing(6)
        self.chat_busy_label = QLabel("")
        self.chat_busy_label.setWordWrap(True)
        self.chat_busy_label.setStyleSheet(
            "color: #9a3412; font-size: 12px; padding: 2px 0;"
        )
        chat_busy_layout.addWidget(self.chat_busy_label, stretch=1)
        self.btn_ai_cancel = QPushButton("중단")
        self.btn_ai_cancel.setToolTip("서버 과부하 대기/재시도를 중단하고 UI를 해제합니다.")
        self.btn_ai_cancel.clicked.connect(self._cancel_ai_request)
        chat_busy_layout.addWidget(self.btn_ai_cancel)
        self.chat_busy_bar.hide()
        right_layout.addWidget(self.chat_busy_bar)

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

    def _on_api_key_edited(self, _text: str = "") -> None:
        """키 변경 시 연결 상태 무효화 (세션 메모리만 유지, 디스크 저장 없음)."""
        if self._api_connected:
            self._api_connected = False
            self._update_api_badge(False, "재연결 필요")

    def _update_api_badge(self, ok: bool, message: str = "") -> None:
        if not hasattr(self, "api_badge"):
            return
        if ok:
            # 메시지에 모델명이 있으면 배지에 함께 표시
            label = "🟢 연결됨"
            msg = (message or "").strip()
            if msg.startswith("연결됨"):
                label = f"🟢 {msg}"
            elif msg:
                label = f"🟢 연결됨 ({msg})"
            if len(label) > 42:
                label = label[:39] + "…)"
            self.api_badge.setText(label)
            self.api_badge.setStyleSheet(
                "QLabel#apiBadge { color: #15803d; background: #dcfce7; }"
            )
            self.api_badge.setToolTip(message or "Gemini API 연결됨")
        else:
            self.api_badge.setText("🔴 미연결")
            self.api_badge.setStyleSheet(
                "QLabel#apiBadge { color: #b91c1c; background: #fee2e2; }"
            )
            self.api_badge.setToolTip(
                message or "STEP 1에서 API Key 입력 후 연결 테스트를 완료해 주세요."
            )

    def _require_api_ready(self) -> bool:
        """분석/챗봇 전 STEP 1 API 연결 확인."""
        key = self.api_key_input.text().strip()
        if key and self._api_connected:
            return True
        self._set_settings_collapsed(False)
        self.api_key_input.setFocus()
        tip = "STEP 1의 API Key 연결이 먼저 필요합니다."
        if not key:
            tip += "\nAPI Key를 입력한 뒤 [연결 테스트]를 눌러 주세요."
        elif not self._api_connected:
            tip += "\n[연결 테스트]가 성공해야 분석을 시작할 수 있습니다."
        self.api_badge.setToolTip(tip)
        self.statusBar().showMessage(tip.replace("\n", " "), 8000)
        # offscreen(자동 테스트)에서는 모달 팝업을 띄우지 않음
        if os.environ.get("QT_QPA_PLATFORM", "").lower() != "offscreen":
            QMessageBox.information(self, "STEP 1 필요", tip)
        return False

    def _set_api_status(self, ok: bool, message: str) -> None:
        self._api_connected = bool(ok)
        self._update_api_badge(ok, message)
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
            QMessageBox.warning(
                self,
                "STEP 1 필요",
                "STEP 1에서 API Key를 입력한 뒤 연결 테스트를 진행해 주세요.",
            )
            return
        self.btn_test.setEnabled(False)
        self.btn_test.setText("확인 중...")
        self.status_api.setText("API: 연결 확인 중...")
        self.status_api.setStyleSheet("color: #64748b;")
        self.api_badge.setText("🟡 확인 중")
        self.api_badge.setStyleSheet(
            "QLabel#apiBadge { color: #a16207; background: #fef9c3; }"
        )
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
        cleaned = normalize_excel_path(path)
        if cleaned:
            self._load_compendium(cleaned)

    def _on_compendium_dropped(self, paths: list[str]) -> None:
        cleaned = [p for p in (normalize_excel_path(x) for x in (paths or [])) if p]
        if cleaned:
            self._load_compendium(cleaned[0])
        else:
            QMessageBox.warning(
                self,
                "공정서 DB",
                "유효한 엑셀 파일(.xlsx / .xls)이 없습니다.",
            )

    def _clear_compendium(self) -> None:
        self.compendium_df = None
        self.compendium_meta = None
        self.compendium_status.setText("공정서 DB: 미등록")
        self.compendium_status.setStyleSheet("color: #64748b; font-size: 12px;")
        self.compendium_status.setToolTip("")
        if self.inventory_data is not None:
            self.inventory_data["compendium_match"] = {}
            self.inventory_data["compendium_match_report"] = ""
        self._update_analyze_button()

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
        self._update_analyze_button()
        # 이미 재고가 로드된 상태면 분석 시작 안내
        if self.inventory_data is not None and os.environ.get("QT_QPA_PLATFORM") != "offscreen":
            QMessageBox.information(
                self,
                "공정서 DB 등록",
                "공정서 DB가 등록되었습니다.\n"
                "[🚀 분석 시작]을 눌러 AI 분석을 실행하세요.",
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
        # AI/챗봇 스냅샷과 동기화
        flags = self.inventory_data.get("ai_flags")
        if isinstance(flags, dict) or flags is None:
            self.inventory_data["ai_flags"] = attach_compendium_match_to_flags(flags, match)
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
        if self.compendium_df is not None:
            name = (self.compendium_meta or {}).get("file_name") or "등록됨"
            self.compendium_status.setText(f"공정서 DB: {name} (이전 데이터 유지)")
            self.compendium_status.setStyleSheet("color: #15803d; font-size: 12px; font-weight: 600;")
        else:
            self.compendium_status.setText("공정서 DB: 미등록")
            self.compendium_status.setStyleSheet("color: #64748b; font-size: 12px;")
        self._update_analyze_button()
        QMessageBox.warning(self, "공정서 DB 오류", message)

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
        cleaned = [p for p in (normalize_excel_path(x) for x in (paths or [])) if p]
        if cleaned:
            self._load_excels(cleaned)

    def _clear_loaded_files(self) -> None:
        self._close_busy()
        self._set_ai_progress(None)
        if self.excel_worker is not None:
            try:
                self.excel_worker.error.disconnect(self._on_excel_error)
                self.excel_worker.finished.disconnect(self._on_excel_loaded)
            except (TypeError, RuntimeError):
                pass
            self.excel_worker = None
        self._loaded_paths = []
        self.inventory_data = None
        self._chat_history.clear()
        self._initial_report = ""
        if hasattr(self, "btn_dl_word"):
            self.btn_dl_word.setEnabled(False)
            self.btn_dl_pdf.setEnabled(False)
        self._report_sections = []
        self._report_section_key = "all"
        self._report_expanded_ids.clear()
        self._chat_busy = False
        self._set_chat_enabled(False)
        self.report_fixed.clear()
        self._clear_report_nav()
        self.chat_view.clear()
        self.table.clear()
        self.table.setRowCount(0)
        self.table.setColumnCount(0)
        self.status_excel.setText("파일: 미업로드")
        self.status_excel.setToolTip("")
        self.status_correction.setText("소급 보정: 0건")
        self.chart._show_placeholder("품목 또는 표준품구분을 선택하면 재고 추이 차트가 표시됩니다.")
        self.chart3d.show_message("엑셀을 다시 업로드해 주세요.")
        self._set_settings_collapsed(False)
        self._update_analyze_button()
        self.statusBar().showMessage("파일 목록이 초기화되었습니다.", 4000)

    def _load_excel(self, file_path: str) -> None:
        self._load_excels([file_path])

    def _restore_excel_ready_status(self) -> None:
        """오류 후에도 UI를 Ready로 유지 (이전 성공 데이터 있으면 복원)."""
        data = self.inventory_data
        if data:
            names = data.get("file_names") or [data.get("file_name", "")]
            if len(names) <= 2:
                file_label = ", ".join(str(n) for n in names if n)
            else:
                file_label = f"{names[0]} 외 {len(names) - 1}개"
            row_n = data.get("row_count")
            suffix = f" ({row_n}품목)" if row_n is not None else ""
            self.status_excel.setText(f"파일: {file_label}{suffix}")
            self.status_excel.setToolTip("\n".join(str(n) for n in names if n))
            self._loaded_paths = list(data.get("file_paths") or self._loaded_paths)
        else:
            self._loaded_paths = []
            self.status_excel.setText("파일: 미업로드")
            self.status_excel.setToolTip("")
        self._update_analyze_button()

    def _load_excels(self, file_paths: list[str]) -> None:
        cleaned = [p for p in (normalize_excel_path(x) for x in (file_paths or [])) if p]
        if not cleaned:
            QMessageBox.warning(
                self,
                "파일 선택",
                "유효한 엑셀 파일(.xlsx / .xls)이 없습니다.\n"
                "(임시 파일 ~$…, 0바이트, 손상 경로는 자동 제외됩니다.)",
            )
            self._restore_excel_ready_status()
            return
        # 순차 업로드: 기존 목록에 합치고 중복 제거
        combined = list(dict.fromkeys(self._loaded_paths + cleaned))
        # 실패 시 롤백을 위해 시도 목록만 임시 보관 (성공 시에만 확정)
        self._pending_excel_paths = list(combined)
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
        # 실패 파일을 pending에서 discard — 이전 성공 목록만 유지
        pending = getattr(self, "_pending_excel_paths", None) or []
        good = set((self.inventory_data or {}).get("file_paths") or [])
        failed_names = [
            Path(p).name for p in pending if p not in good
        ] or [Path(p).name for p in pending]
        self._pending_excel_paths = []
        self._restore_excel_ready_status()
        detail = "\n".join(f"- {n}" for n in failed_names[:12])
        if len(failed_names) > 12:
            detail += f"\n- … 외 {len(failed_names) - 12}개"
        QMessageBox.warning(
            self,
            "파일 처리 오류",
            f"일부/전체 엑셀 로드에 실패했습니다.\n"
            f"실패한 파일은 목록에서 제외했고, 시스템은 사용 가능한 상태입니다.\n\n"
            f"{message}\n\n{detail}".strip(),
        )

    def _on_excel_loaded(self, data: dict[str, Any]) -> None:
        self._close_busy()
        # 성공 경로만 목록에 확정 (실패 파일은 discard)
        loaded = list(data.get("file_paths") or [])
        self._loaded_paths = loaded
        self._pending_excel_paths = []
        failed = list(data.get("failed_files") or [])
        stock_items = list(data.get("stock_items") or [])
        data["ai_flags"] = collect_ai_analysis_flags(stock_items)
        data["ai_mentioned_codes"] = []
        self.inventory_data = data
        self._refresh_compendium_match()
        self._chat_history.clear()
        self._initial_report = ""
        if hasattr(self, "btn_dl_word"):
            self.btn_dl_word.setEnabled(False)
            self.btn_dl_pdf.setEnabled(False)
        self._report_sections = []
        self._report_section_key = "all"
        self._report_expanded_ids.clear()
        self._set_chat_enabled(False)
        self.report_fixed.clear()
        self._clear_report_nav()
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
        stats = match.get("stats") or {}
        miss_n = stats.get("missing_count")
        match_n = match.get("correction_count")
        match_bits = []
        if match_n is not None and match:
            match_bits.append(f"매칭 {match_n}")
        if miss_n is not None and match:
            match_bits.append(f"미보유 {miss_n}")
        match_txt = f" · 공정서 {'/'.join(match_bits)}" if match_bits else ""
        self.status_excel.setText(f"파일: {file_label} ({data['row_count']}품목)")
        self.status_excel.setToolTip("\n".join(names))
        self.status_correction.setText(
            f"소급 보정: {data['correction_count']:,}건 · 변동 {changed}건 · "
            f"소진후보 {deplete_n} · 가속 {surge_n}{val_txt}{match_txt}"
        )
        self._populate_table(data)
        self._populate_filters(data)
        self._refresh_3d()
        self._update_analyze_button()
        self.tabs.setCurrentIndex(0)
        # 파일 로드만 수행 — AI 분석은 [🚀 분석 시작]에서 실행
        msg = "재고 파일 로드 완료. 공정서 DB 등록 후 [🚀 분석 시작]을 눌러 주세요."
        if failed:
            fail_txt = "\n".join(
                f"- {f.get('name')}: {f.get('error')}" for f in failed[:8]
            )
            if len(failed) > 8:
                fail_txt += f"\n- … 외 {len(failed) - 8}개"
            QMessageBox.warning(
                self,
                "일부 파일 로드 실패",
                f"{len(failed)}개 파일은 제외하고 나머지 {len(loaded)}개를 로드했습니다.\n\n"
                f"{fail_txt}",
            )
            msg = f"일부 파일 제외 후 로드 완료 ({len(loaded)}개 성공 · {len(failed)}개 실패)."
        self.statusBar().showMessage(msg, 8000)

    def _update_analyze_button(self) -> None:
        has_stock = self.inventory_data is not None
        has_comp = self.compendium_df is not None
        ready = has_stock and has_comp
        if hasattr(self, "btn_analyze"):
            self.btn_analyze.setEnabled(ready)
        if hasattr(self, "analyze_hint"):
            if ready:
                self.analyze_hint.setText(
                    "준비 완료 — [🚀 분석 시작]을 누르면 AI 분석이 실행되고 상단이 접힙니다."
                )
                self.analyze_hint.setStyleSheet("color: #0f766e; font-size: 11px; font-weight: 600;")
            elif has_stock and not has_comp:
                self.analyze_hint.setText("공정서 DB를 등록하면 분석을 시작할 수 있습니다.")
                self.analyze_hint.setStyleSheet("color: #64748b; font-size: 11px;")
            elif has_comp and not has_stock:
                self.analyze_hint.setText("재고 엑셀을 등록하면 분석을 시작할 수 있습니다.")
                self.analyze_hint.setStyleSheet("color: #64748b; font-size: 11px;")
            else:
                self.analyze_hint.setText(
                    "재고 엑셀 + 공정서 DB 등록 후 [🚀 분석 시작]을 누르면 AI 분석이 실행됩니다."
                )
                self.analyze_hint.setStyleSheet("color: #64748b; font-size: 11px;")

    def _on_analyze_clicked(self) -> None:
        if not self.inventory_data:
            tip = "STEP 2에서 재고 엑셀을 먼저 등록해 주세요."
            if os.environ.get("QT_QPA_PLATFORM") != "offscreen":
                QMessageBox.information(self, "분석 시작", tip)
            return
        if self.compendium_df is None:
            tip = "STEP 2에서 공정서 DB를 먼저 등록해 주세요."
            self._set_settings_collapsed(False)
            if os.environ.get("QT_QPA_PLATFORM") != "offscreen":
                QMessageBox.information(self, "분석 시작", tip)
            return
        if not self._require_api_ready():
            return
        self._set_settings_collapsed(True)
        self.tabs.setCurrentIndex(2)  # AI 분석 탭
        self._run_ai_analysis(self.inventory_data)

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
                    orig_text = "" if orig is None else format_qty_int(orig)
                    corr_text = "" if corr is None else format_qty_int(corr)
                    if orig_text == "-":
                        orig_text = ""
                    if corr_text == "-":
                        corr_text = ""
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

    def _focus_table_row(self, row: int) -> None:
        """분양현황 테이블에서 행 선택·스크롤·포커스."""
        if row < 0 or row >= self.table.rowCount():
            return
        self.table.clearSelection()
        self.table.selectRow(row)
        self.table.setCurrentCell(row, 0)
        item = self.table.item(row, 0)
        if item is not None:
            self.table.scrollToItem(
                item, self.table.ScrollHint.PositionAtCenter
            )
        self.table.setFocus(Qt.FocusReason.OtherFocusReason)

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
        """3D 산점도 점 클릭 → 분양현황 탭으로 전환 후 해당 행 스크롤·선택."""
        if not self.inventory_data:
            return
        items = list(self.inventory_data.get("items") or [])
        target = None
        target_row = -1
        code = (code or "").strip()
        name = (name or "").strip()
        if code:
            for idx, it in enumerate(items):
                mgmt = str(it.get("mgmt_no") or it.get("manage_no") or "")
                label = str(it.get("label") or "")
                if mgmt == code or label.startswith(f"{code}(") or code in label:
                    target = it
                    target_row = idx
                    break
        if target is None and name:
            for idx, it in enumerate(items):
                label = str(it.get("label") or "")
                if name in label or label.endswith(f"({name})"):
                    target = it
                    target_row = idx
                    break
        # 분양현황 탭(0)으로 전환 — 분양차트 이동 중단
        self.tabs.setCurrentIndex(0)
        if target is None:
            return
        row = target_row
        if row < 0 or row >= self.table.rowCount():
            # 라벨로 테이블 행 재검색
            want = str(target.get("label") or "")
            for r in range(self.table.rowCount()):
                cell = self.table.item(r, 0)
                stored = ""
                if cell is not None:
                    stored = str(cell.data(Qt.ItemDataRole.UserRole) or "")
                if stored == want or (want and want in stored):
                    row = r
                    break
        if 0 <= row < self.table.rowCount():
            self._focus_table_row(row)

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

    def _on_global_find_shortcut(self) -> None:
        idx = self.tabs.currentIndex()
        if idx == 0:
            self._show_table_find_bar()
        elif idx == 2:
            self._show_report_find_bar()

    def _show_table_find_bar(self) -> None:
        self.table_find_bar.show()
        self.table_find_input.setFocus()
        self.table_find_input.selectAll()

    def _hide_table_find_bar(self) -> None:
        self.table_find_bar.hide()
        self.table.clearSelection()
        self.table.setFocus()

    def _table_searchable_columns(self) -> list[int]:
        cols: list[int] = []
        for i in range(self.table.columnCount()):
            header_item = self.table.horizontalHeaderItem(i)
            header = header_item.text() if header_item is not None else ""
            if any(k in header for k in ("관리번호", "한글", "영문", "성분", "표준품")):
                cols.append(i)
        if cols:
            return cols
        meta_n = len((self.inventory_data or {}).get("meta_cols") or [])
        return list(range(min(meta_n, self.table.columnCount())))

    def _rebuild_table_find_matches(self, query: str) -> None:
        q = query.strip().lower()
        self._table_find_matches = []
        self._table_find_index = -1
        if not q:
            return
        for row in range(self.table.rowCount()):
            for col in self._table_searchable_columns():
                item = self.table.item(row, col)
                if item is not None and q in item.text().lower():
                    self._table_find_matches.append((row, col))

    def _focus_table_match(self, row: int, col: int) -> None:
        self.table.setCurrentCell(row, col)
        self.table.selectRow(row)
        item = self.table.item(row, col)
        if item is not None:
            self.table.scrollToItem(item)

    def _find_in_table(self, forward: bool = True) -> None:
        query = self.table_find_input.text()
        if not query.strip():
            return
        if not self._table_find_matches or query.strip().lower() != getattr(
            self, "_table_find_last_query", ""
        ):
            self._rebuild_table_find_matches(query)
            self._table_find_last_query = query.strip().lower()
        if not self._table_find_matches:
            self.statusBar().showMessage(f"검색 결과 없음: {query}", 3000)
            return
        if forward:
            self._table_find_index = (self._table_find_index + 1) % len(self._table_find_matches)
        else:
            self._table_find_index = (
                self._table_find_index - 1
            ) % len(self._table_find_matches)
        row, col = self._table_find_matches[self._table_find_index]
        self._focus_table_match(row, col)
        self.statusBar().showMessage(
            f"검색 {self._table_find_index + 1}/{len(self._table_find_matches)}: {query}",
            4000,
        )

    def _download_report_word(self) -> None:
        """AI 리포트를 Word(.docx)로 저장 — 본문 검사·경량 변환·WaitCursor."""
        from PyQt6.QtWidgets import QFileDialog

        placeholder_hints = (
            "생약표준품 분양·소진 예측 AI 리포트가 여기에 표시됩니다",
            "엑셀 업로드 후",
        )
        viewer_text = ""
        if hasattr(self, "report_fixed"):
            try:
                viewer_text = self.report_fixed.toPlainText().strip()
            except Exception:
                viewer_text = ""
        report_text = (self._initial_report or "").strip() or viewer_text
        if (
            not report_text
            or any(h in report_text for h in placeholder_hints)
            or len(report_text) < 40
        ):
            QMessageBox.warning(
                self,
                "안내",
                "내보낼 AI 분석 리포트가 없습니다. 먼저 분석을 완료해 주세요.",
            )
            return

        path, _ = QFileDialog.getSaveFileName(
            self,
            "Word 리포트 저장",
            "생약표준품_AI_분석리포트.docx",
            "Word Files (*.docx)",
        )
        if not path:
            return
        if not path.lower().endswith(".docx"):
            path += ".docx"

        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            export_markdown_report_to_docx(report_text, path)
            QApplication.processEvents()
            QMessageBox.information(self, "완료", "Word 파일 저장이 완료되었습니다.")
            self.statusBar().showMessage(f"Word 저장 완료: {path}", 5000)
        except Exception as exc:
            QMessageBox.critical(
                self, "저장 오류", f"Word 파일 생성 실패: {exc}"
            )
        finally:
            QApplication.restoreOverrideCursor()

    def _download_report_pdf(self) -> None:
        """현재 렌더링된 리포트 서식을 PDF로 저장 (QPdfWriter)."""
        from PyQt6.QtWidgets import QFileDialog

        if not self._initial_report:
            QMessageBox.information(self, "안내", "먼저 AI 분석 리포트를 생성해 주세요.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self,
            "PDF 저장",
            "생약표준품_AI_분석_리포트.pdf",
            "PDF (*.pdf)",
        )
        if not path:
            return
        if not path.lower().endswith(".pdf"):
            path += ".pdf"
        try:
            writer = QPdfWriter(path)
            writer.setTitle("생약표준품 AI 분석 리포트")
            writer.setPageLayout(
                QPageLayout(
                    QPageSize(QPageSize.PageSizeId.A4),
                    QPageLayout.Orientation.Portrait,
                    QMarginsF(12, 12, 12, 12),
                    QPageLayout.Unit.Millimeter,
                )
            )
            html = markdown_report_to_collapsible_html(
                self._initial_report,
                expanded_ids=set(self._report_expanded_ids),
            )
            doc = QTextDocument()
            doc.setDefaultFont(QFont("Malgun Gothic", 10))
            doc.setHtml(html)
            doc.print(writer)
            self.statusBar().showMessage(f"PDF 저장 완료: {path}", 5000)
        except Exception as exc:
            QMessageBox.critical(self, "PDF 저장 실패", str(exc))

    def _show_report_find_bar(self) -> None:
        self.report_find_bar.show()
        self.report_find_input.setFocus()
        self.report_find_input.selectAll()

    def _hide_report_find_bar(self) -> None:
        self.report_find_bar.hide()
        self.report_fixed.setExtraSelections([])
        self.report_fixed.setFocus()

    def _vertical_nav_label(self, text: str) -> str:
        """사이드 버튼 라벨 — 가독성을 위해 가로 표기(미보유·검토 등)."""
        label = (text or "").strip() or "·"
        return label

    def _clear_report_nav(self) -> None:
        for btn in list(self._report_nav_buttons.values()):
            self._report_nav_group.removeButton(btn)
            btn.deleteLater()
        self._report_nav_buttons.clear()
        while self.report_nav_layout.count():
            item = self.report_nav_layout.takeAt(0)
            w = item.widget()
            if w is not None:
                w.deleteLater()
        self.report_nav_layout.addStretch(1)

    def _rebuild_report_nav(self) -> None:
        """고정 섹션 버튼(미보유·검토 포함) 재구성."""
        self._clear_report_nav()
        if not self._report_sections and not self._initial_report:
            return

        while self.report_nav_layout.count():
            self.report_nav_layout.takeAt(0)

        entries: list[tuple[str, str, str]] = [("all", "전체", "전체 리포트")]
        for sec in self._report_sections:
            entries.append(
                (sec["id"], sec.get("short") or "항목", sec.get("title") or "")
            )

        for i, (key, short, tip) in enumerate(entries):
            btn = QPushButton(self._vertical_nav_label(short))
            btn.setObjectName("reportNavBtn")
            btn.setCheckable(True)
            btn.setToolTip(tip or short)
            btn.setProperty("section_key", key)
            btn.setMinimumHeight(40)
            self._report_nav_group.addButton(btn, i)
            self._report_nav_buttons[key] = btn
            self.report_nav_layout.addWidget(btn)

        self.report_nav_layout.addStretch(1)

        default_key = "summary" if "summary" in self._report_nav_buttons else (
            self._report_sections[0]["id"] if self._report_sections else "all"
        )
        if self._report_section_key not in self._report_nav_buttons:
            self._report_section_key = default_key
        btn = self._report_nav_buttons.get(self._report_section_key)
        if btn is not None:
            btn.setChecked(True)

    def _on_report_nav_clicked(self, btn_id: int) -> None:
        btn = self._report_nav_group.button(btn_id)
        if btn is None:
            return
        key = btn.property("section_key")
        if key:
            self._report_section_key = str(key)
            self._render_report_html()

    def _report_context_for_sections(self) -> tuple[dict, dict | None]:
        data = self.inventory_data or {}
        flags = dict(data.get("ai_flags") or {})
        match = data.get("compendium_match")
        if match:
            flags = attach_compendium_match_to_flags(flags, match)
        return flags, match

    def _render_report_html(self) -> None:
        """표준 리포트를 섹션 버튼 선택에 맞춰 HTML로 표시."""
        if not self._initial_report:
            self.report_fixed.clear()
            self._clear_report_nav()
            self._report_sections = []
            return

        flags, match = self._report_context_for_sections()
        if not self._report_sections:
            self._report_sections = split_markdown_report_sections(
                self._initial_report, flags=flags, match_result=match
            )
            self._rebuild_report_nav()

        md = self._initial_report
        if self._report_section_key != "all":
            for sec in self._report_sections:
                if sec["id"] == self._report_section_key:
                    md = sec["markdown"]
                    break

        self.report_fixed.setHtml(
            markdown_report_to_collapsible_html(
                md,
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

    def _set_ai_progress(self, message: str | None) -> None:
        """모달 없이 챗봇 인라인 + 상태바로 AI 진행 표시 (메인 UI 조작 가능)."""
        if not hasattr(self, "chat_busy_label"):
            return
        if message:
            self.chat_busy_label.setText(f"🤖 {message}")
            self.chat_busy_bar.show()
            if hasattr(self, "btn_ai_cancel"):
                self.btn_ai_cancel.setEnabled(True)
            self.statusBar().showMessage(message)
        else:
            self.chat_busy_bar.hide()
            self.chat_busy_label.clear()
            if hasattr(self, "btn_ai_cancel"):
                self.btn_ai_cancel.setEnabled(False)
            self.statusBar().clearMessage()

    def _cancel_ai_request(self) -> None:
        """과부하 대기/재시도 중 AI 요청을 중단하고 UI를 해제."""
        if not self._chat_busy:
            return
        request_cancel_gemini()
        self._set_ai_progress("취소 요청됨 — 진행 중인 대기를 중단합니다…")
        if hasattr(self, "btn_ai_cancel"):
            self.btn_ai_cancel.setEnabled(False)

    def _on_ai_stage(self, message: str) -> None:
        self._set_ai_progress(message)
        # 챗봇 시스템 안내 문구도 단계에 맞게 갱신
        for msg in reversed(self._chat_history):
            if msg.get("role") == "system":
                msg["text"] = message
                self._render_chat_view()
                break

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
        match = data.get("compendium_match") or {}
        data["ai_flags"] = attach_compendium_match_to_flags(data.get("ai_flags"), match)
        self._chat_history.clear()
        self._initial_report = ""
        if hasattr(self, "btn_dl_word"):
            self.btn_dl_word.setEnabled(False)
            self.btn_dl_pdf.setEnabled(False)
        self._report_sections = []
        self._report_section_key = "all"
        self._report_expanded_ids.clear()
        self.report_fixed.clear()
        self._clear_report_nav()
        self._set_chat_enabled(False)
        if not changed:
            self.chat_view.setMarkdown(
                "분석 대상 없음\n\n수량 변화가 있는 품목이 없어 AI 분석을 건너뛰었습니다."
            )
            self._scroll_chat_to_bottom()
            return
        if not self._require_api_ready():
            self.chat_view.setMarkdown(
                "**STEP 1의 API Key 연결이 먼저 필요합니다.**\n\n"
                "상단에서 Gemini API Key를 입력하고 [연결 테스트]를 완료한 뒤, "
                "[🚀 분석 시작]을 눌러 주세요.\n\n"
                f"(대상 품목 중 변동 {len(changed)}건 대기 중)"
            )
            self._scroll_chat_to_bottom()
            return
        key = self.api_key_input.text().strip()
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
                    f"소진후보 {deplete_n} · 가속 {surge_n} · 제조검토후보 {mfg_n})"
                ),
            }
        )
        self._render_chat_view()
        self._set_ai_progress("재고 및 공정서 DB 분석 중...")

        stock_items_snap = list(stock_items)
        flags_snap = data.get("ai_flags")
        match_report = data.get("compendium_match_report") or None
        compendium_df = self.compendium_df
        compendium_meta = self.compendium_meta

        def _build_prompt() -> str:
            ctx = None
            if compendium_df is not None:
                ctx = format_compendium_context(
                    compendium_df,
                    stock_items_snap,
                    meta=compendium_meta,
                ) or None
            return build_ai_prompt(
                stock_items_snap,
                compendium_context=ctx,
                compendium_match_report=match_report,
                flags=flags_snap,
            )

        self.gemini_worker = GeminiWorker(
            key, build_prompt=_build_prompt, followup=False
        )
        self.gemini_worker.stage.connect(self._on_ai_stage)
        self.gemini_worker.finished.connect(self._on_report_ready)
        self.gemini_worker.error.connect(self._on_report_error)
        self.gemini_worker.start()

    def _on_report_ready(self, text: str) -> None:
        self._set_ai_progress(None)
        self._chat_busy = False
        clear_cancel_gemini()
        self._chat_history.clear()
        flags, match = self._report_context_for_sections()
        text = ensure_mandatory_report_sections(text, flags=flags, match_result=match)
        self._initial_report = text
        self._report_expanded_ids.clear()
        self._report_sections = split_markdown_report_sections(
            text, flags=flags, match_result=match
        )
        self._report_section_key = (
            "summary"
            if any(s["id"] == "summary" for s in self._report_sections)
            else (self._report_sections[0]["id"] if self._report_sections else "all")
        )
        self._rebuild_report_nav()
        self._render_report_html()
        self.btn_dl_word.setEnabled(True)
        self.btn_dl_pdf.setEnabled(True)
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
        self._set_ai_progress(None)
        self._chat_busy = False
        clear_cancel_gemini()
        self._chat_history = [m for m in self._chat_history if m.get("role") != "system"]
        hint = ""
        low = (message or "").lower()
        if any(k in low for k in ("과부하", "503", "429", "unavailable", "overload")) or "취소" in (message or ""):
            hint = (
                "\n\n💡 **안내:** 서버 과부하 시 앱이 대체 Flash 모델로 자동 전환합니다. "
                "그래도 실패하면 1~2분 뒤 [🚀 분석 시작]을 다시 눌러 주세요. "
                "대기 중에는 [중단]으로 UI를 해제할 수 있습니다."
            )
        self._chat_history.append(
            {"role": "assistant", "text": f"AI 분석 생성 실패:\n\n{message}{hint}"}
        )
        self._render_chat_view()
        self._set_chat_enabled(bool(self._initial_report))
        # 일시 과부하는 API Key 자체 문제로 보지 않음
        if "취소" in (message or "") or any(
            k in low for k in ("과부하", "503", "429", "unavailable", "overload")
        ):
            self._set_api_status(True, f"연결됨 ({get_active_gemini_model()}) — 일시 과부하")
        else:
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
        if not self._require_api_ready():
            return
        key = self.api_key_input.text().strip()

        self.chat_input.clear()
        self._chat_history.append({"role": "user", "text": question})
        self._chat_history.append(
            {"role": "system", "text": "재고 및 공정서 DB 분석 중..."}
        )
        self._chat_busy = True
        self._set_chat_enabled(False)
        self._render_chat_view()
        self._set_ai_progress("재고 및 공정서 DB 분석 중...")

        # 메인 스레드에서 참조만 스냅샷 — 무거운 프롬프트 구성은 워커에서 수행
        base_report = self._initial_report
        stock_items = None
        table_df = None
        ai_flags = None
        if self.inventory_data:
            stock_items = self.inventory_data.get("stock_items")
            table_df = self.inventory_data.get("table_df")
            ai_flags = attach_compendium_match_to_flags(
                self.inventory_data.get("ai_flags"),
                self.inventory_data.get("compendium_match"),
            )
            self.inventory_data["ai_flags"] = ai_flags
        compendium_df = self.compendium_df
        compendium_meta = self.compendium_meta
        prior_qas = [
            m for m in self._chat_history
            if m.get("role") in ("user", "assistant")
        ]

        def _build_prompt() -> str:
            ctx = None
            if compendium_df is not None:
                ctx = format_compendium_context(
                    compendium_df,
                    stock_items or [],
                    meta=compendium_meta,
                ) or None
            prompt = build_followup_prompt(
                base_report,
                question,
                stock_items,
                compendium_context=ctx,
                table_df=table_df,
                flags=ai_flags,
            )
            if len(prior_qas) > 2:
                tail = prior_qas[-6:-1]
                if tail:
                    ctx_chat = "\n\n".join(
                        ("사용자: " if m["role"] == "user" else "AI: ") + m["text"][:800]
                        for m in tail
                    )
                    prompt = prompt + "\n\n[최근 대화 맥락]\n" + ctx_chat
            return prompt

        self.gemini_worker = GeminiWorker(
            key, build_prompt=_build_prompt, followup=True
        )
        self.gemini_worker.stage.connect(self._on_ai_stage)
        self.gemini_worker.finished.connect(self._on_chat_reply)
        self.gemini_worker.error.connect(self._on_chat_error)
        self.gemini_worker.start()

    def _on_chat_reply(self, text: str) -> None:
        self._set_ai_progress(None)
        self._chat_busy = False
        clear_cancel_gemini()
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
        self._set_ai_progress(None)
        self._chat_busy = False
        clear_cancel_gemini()
        self._chat_history = [m for m in self._chat_history if m.get("role") != "system"]
        hint = ""
        if any(
            k in (message or "").lower()
            for k in ("과부하", "503", "429", "unavailable", "overload")
        ) or "취소" in (message or ""):
            hint = (
                "\n\n💡 잠시 후 다시 질문해 주세요. "
                "앱은 과부하 시 대체 모델로 자동 전환을 시도합니다."
            )
        self._chat_history.append(
            {"role": "assistant", "text": f"답변 생성 실패:\n\n{message}{hint}"}
        )
        self._render_chat_view()
        self._set_chat_enabled(True)
        low = (message or "").lower()
        if "취소" in (message or "") or any(
            k in low for k in ("과부하", "503", "429", "unavailable", "overload")
        ):
            self._set_api_status(True, f"연결됨 ({get_active_gemini_model()}) — 일시 과부하")
        else:
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
