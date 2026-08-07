"""
재고 분석 및 소급 보정 프로그램 (PyQt6)
- 엑셀 양식: 기본정보(A~H) + 변동기록(I열~ 변경일자/재고량 쌍)
"""

from __future__ import annotations

import json
import sys
import traceback
from datetime import datetime, date
from pathlib import Path
from typing import Any

import matplotlib
import pandas as pd
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from PyQt6.QtCore import QEvent, QObject, Qt, QThread, pyqtSignal
from PyQt6.QtGui import QColor, QDragEnterEvent, QDropEvent, QFont
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
    QPushButton,
    QStatusBar,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

matplotlib.use("QtAgg")
matplotlib.rcParams["font.family"] = "Malgun Gothic"
matplotlib.rcParams["axes.unicode_minus"] = False

CONFIG_PATH = Path(__file__).parent / "config.json"
META_COL_COUNT = 8  # A~H: 순번~분양여부
INVENTORY_START_COL = 8  # I열 (0-based)

# 응답이 빠른 모델 우선. (3.5/3.6은 응답 지연이 길어 연결 테스트가 멈춘 것처럼 보일 수 있음)
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
QPushButton:hover {
    background-color: #274b7a;
}
QPushButton:pressed {
    background-color: #152a45;
}
QPushButton#secondaryBtn {
    background-color: #ffffff;
    color: #1e3a5f;
    border: 1px solid #1e3a5f;
}
QPushButton#secondaryBtn:hover {
    background-color: #eef3f9;
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
    padding: 9px 18px;
    margin-right: 4px;
    min-width: 110px;
}
QTabBar::tab:selected {
    background: #ffffff;
    color: #0b1f3a;
    font-weight: 700;
    border-color: #d8e0ea;
}
QTabBar::tab:hover:!selected {
    background: #dce6f2;
}
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
QStatusBar QLabel {
    color: #334155;
    padding: 2px 10px;
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
    """터미널(stderr)에 Gemini 관련 로그 출력."""
    print(f"[Gemini {level}] {message}", file=sys.stderr, flush=True)
    if exc is not None:
        print(traceback.format_exc(), file=sys.stderr, flush=True)


def format_gemini_error(exc: BaseException) -> str:
    """google-genai 및 기타 예외를 사용자 친화적 메시지로 변환."""
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
    """목록엔 있으나 신규 사용자에게 차단/폐기된 모델 오류인지 판별."""
    text = str(exc).lower()
    markers = (
        "no longer available",
        "not found",
        "not_found",
        "is not found",
        "unsupported",
        "resource_exhausted",
        "rate limit",
        "quota",
        "deadline",
        "timed out",
        "timeout",
    )
    if any(m in text for m in markers):
        return True
    code = _api_error_code(exc)
    if code in (404, 400, 429, 503, 504):
        return True
    try:
        from google.genai import errors as genai_errors

        if isinstance(exc, genai_errors.APIError):
            status = str(getattr(exc, "status", "") or "").upper()
            return status in {
                "NOT_FOUND",
                "INVALID_ARGUMENT",
                "RESOURCE_EXHAUSTED",
                "UNAVAILABLE",
                "DEADLINE_EXCEEDED",
            }
    except ImportError:
        pass
    return False


def _candidate_models(prioritize: str | None = None) -> list[str]:
    """실제 호출 검증할 후보(소수)만 구성. 전체 모델 목록 조회는 하지 않음."""
    ordered: list[str] = []
    if prioritize:
        ordered.append(prioritize)
    if _resolved_gemini_model and _resolved_gemini_model not in ordered:
        ordered.append(_resolved_gemini_model)
    for preferred in GEMINI_MODEL_PREFERENCES:
        if preferred not in ordered:
            ordered.append(preferred)
    return ordered[:_MAX_MODEL_PROBES]


def _probe_model(client, model: str) -> bool:
    """짧은 호출로 실제 generateContent 가능 여부 확인."""
    from google.genai import types

    try:
        response = client.models.generate_content(
            model=model,
            contents="Reply with exactly one word: OK",
            config=types.GenerateContentConfig(
                max_output_tokens=16,
                temperature=0,
            ),
        )
        ok = extract_response_text(response) is not None
        if not ok:
            log_gemini("WARN", f"모델 응답 비어 있음, 건너뜀: {model}")
        return ok
    except Exception as exc:
        if _is_auth_error(exc):
            raise
        log_gemini("WARN", f"모델 사용 불가, 건너뜀: {model} ({format_gemini_error(exc)})")
        return False


def resolve_gemini_model(client, force_refresh: bool = False) -> str:
    """실제 호출이 되는 Flash 모델을 빠르게 선택."""
    global _resolved_gemini_model
    if _resolved_gemini_model and not force_refresh:
        return _resolved_gemini_model

    # 연결 테스트/재탐색 시에도 기존 성공 모델을 최우선으로 시도
    candidates = _candidate_models(prioritize=_resolved_gemini_model)
    errors: list[str] = []
    for model in candidates:
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

    raise RuntimeError(
        "사용 가능한 Gemini Flash 모델을 찾지 못했습니다.\n" + "\n".join(errors[:8])
    )


def get_active_gemini_model() -> str:
    return _resolved_gemini_model or GEMINI_MODEL_PREFERENCES[0]


def describe_empty_response(response: Any) -> str:
    """응답은 왔지만 text가 없을 때 원인 설명."""
    parts: list[str] = ["모델 응답 본문이 비어 있습니다."]

    prompt_feedback = getattr(response, "prompt_feedback", None)
    if prompt_feedback:
        block_reason = getattr(prompt_feedback, "block_reason", None)
        if block_reason:
            parts.append(f"프롬프트 차단 사유: {block_reason}")

    candidates = getattr(response, "candidates", None) or []
    if not candidates:
        parts.append("candidates가 없습니다.")
    else:
        candidate = candidates[0]
        finish_reason = getattr(candidate, "finish_reason", None)
        if finish_reason:
            parts.append(f"finish_reason: {finish_reason}")
        safety_ratings = getattr(candidate, "safety_ratings", None)
        if safety_ratings:
            parts.append(f"safety_ratings: {safety_ratings}")

    return " ".join(parts)


def extract_response_text(response: Any) -> str | None:
    text = getattr(response, "text", None)
    if text and str(text).strip():
        return str(text).strip()
    return None


def test_gemini_connection(api_key: str) -> tuple[bool, str]:
    """Gemini API 연결 테스트. (성공 여부, 상세 메시지) 반환."""
    global _resolved_gemini_model
    client = create_gemini_client(api_key)
    log_gemini("INFO", "연결 테스트 시작")

    # 1) 캐시된 모델만 빠르게 확인
    if _resolved_gemini_model:
        log_gemini("INFO", f"캐시 모델 확인: {_resolved_gemini_model}")
        if _probe_model(client, _resolved_gemini_model):
            msg = f"연결됨 ({_resolved_gemini_model})"
            log_gemini("INFO", msg)
            return True, msg
        log_gemini("WARN", f"캐시 모델 실패, 재탐색: {_resolved_gemini_model}")
        _resolved_gemini_model = None

    # 2) 소수 후보만 순차 검증
    model = resolve_gemini_model(client, force_refresh=True)
    msg = f"연결됨 ({model})"
    log_gemini("INFO", msg)
    return True, msg


def generate_gemini_report(api_key: str, prompt: str) -> str:
    global _resolved_gemini_model
    client = create_gemini_client(api_key)
    model = resolve_gemini_model(client)
    log_gemini("INFO", f"AI 분석 요청 시작 (model={model})")

    try:
        response = client.models.generate_content(
            model=model,
            contents=prompt,
        )
    except Exception as exc:
        if _is_model_unavailable_error(exc):
            log_gemini("WARN", f"선택 모델 실패, 재탐색: {model} ({format_gemini_error(exc)})")
            _resolved_gemini_model = None
            model = resolve_gemini_model(client, force_refresh=True)
            response = client.models.generate_content(
                model=model,
                contents=prompt,
            )
        else:
            raise

    text = extract_response_text(response)
    if text:
        log_gemini("INFO", "AI 분석 리포트 생성 완료")
        return text

    detail = describe_empty_response(response)
    log_gemini("ERROR", detail)
    raise RuntimeError(detail)


# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------


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
# Data processing
# ---------------------------------------------------------------------------


def apply_retroactive_correction(values: list[float | None]) -> tuple[list[float | None], int]:
    """재고 증가 시점의 차액만큼 과거 수량을 소급 상향 조정."""
    corrected = list(values)
    correction_count = 0

    for j in range(1, len(corrected)):
        prev_val = corrected[j - 1]
        curr_val = corrected[j]
        if prev_val is None or curr_val is None:
            continue
        if curr_val > prev_val:
            delta = curr_val - prev_val
            for k in range(j):
                if corrected[k] is not None:
                    corrected[k] += delta
                    correction_count += 1

    return corrected, correction_count


def to_float_or_none(val: Any) -> float | None:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, str):
        text = val.strip().replace(",", "")
        if not text:
            return None
        val = text
    try:
        return float(val)
    except (TypeError, ValueError):
        return None


def is_empty(val: Any) -> bool:
    if val is None:
        return True
    try:
        if pd.isna(val):
            return True
    except (TypeError, ValueError):
        pass
    if isinstance(val, str) and not val.strip():
        return True
    return False


def format_date_value(val: Any) -> str | None:
    """변경일자를 YYYY-MM-DD로 정규화. 빈값/NaN/NaT는 None."""
    if is_empty(val):
        return None

    if isinstance(val, pd.Timestamp):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, date):
        return val.strftime("%Y-%m-%d")

    text = str(val).strip()
    if not text or text.lower() in {"nan", "nat", "none"}:
        return None

    parsed = pd.to_datetime(text, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.strftime("%Y-%m-%d")


def find_meta_index(columns: list[Any], candidates: list[str], default: int) -> int:
    normalized = [str(c).strip() for c in columns]
    for name in candidates:
        if name in normalized:
            return normalized.index(name)
    return default


def item_label(mgmt_no: str, korean_name: str) -> str:
    mgmt_no = (mgmt_no or "").strip()
    korean_name = (korean_name or "").strip()
    if mgmt_no and korean_name:
        return f"{mgmt_no}({korean_name})"
    return mgmt_no or korean_name or "(미식별)"


def extract_change_pairs(row: pd.Series, pair_cols: list[Any]) -> tuple[list[str], list[float | None]]:
    """I열 이후 변경일자/재고량 쌍을 추출. 날짜가 없는 쌍은 제외."""
    dates: list[str] = []
    qtys: list[float | None] = []

    for i in range(0, len(pair_cols) - 1, 2):
        date_col = pair_cols[i]
        qty_col = pair_cols[i + 1]
        date_str = format_date_value(row[date_col])
        if date_str is None:
            continue
        dates.append(date_str)
        qtys.append(to_float_or_none(row[qty_col]))

    return dates, qtys


def process_excel(file_path: str) -> dict[str, Any]:
    df = pd.read_excel(file_path, header=0, engine="openpyxl")

    if df.shape[1] <= INVENTORY_START_COL:
        raise ValueError(
            "I열(9번째 열) 이후 변동기록(변경일자/재고량 쌍)이 존재하지 않습니다.\n"
            "기본정보 8열 + 변동기록 쌍 양식인지 확인해 주세요."
        )

    # 기본정보 열 인식 (헤더명 우선, 없으면 위치 기반)
    raw_cols = list(df.columns)
    meta_cols = raw_cols[:META_COL_COUNT]
    pair_cols = raw_cols[INVENTORY_START_COL:]

    # 마지막 유효 데이터 열까지 동적 인식
    last_idx = len(pair_cols) - 1
    for i in range(len(pair_cols) - 1, -1, -1):
        col = pair_cols[i]
        if df[col].notna().any():
            last_idx = i
            break
    pair_cols = pair_cols[: last_idx + 1]

    if len(pair_cols) < 2:
        raise ValueError("변경일자/재고량 쌍 데이터를 찾을 수 없습니다.")

    # 홀수 개면 마지막 불완전 열 제외
    if len(pair_cols) % 2 == 1:
        pair_cols = pair_cols[:-1]

    mgmt_idx = find_meta_index(meta_cols, ["관리번호"], 2)
    name_idx = find_meta_index(meta_cols, ["한글명"], 3)

    items: list[dict[str, Any]] = []
    total_corrections = 0
    max_pair_count = 0

    for row_i, (_, row) in enumerate(df.iterrows()):
        meta_vals = ["" if is_empty(row[c]) else str(row[c]).strip() for c in meta_cols]
        mgmt_no = meta_vals[mgmt_idx] if mgmt_idx < len(meta_vals) else ""
        korean_name = meta_vals[name_idx] if name_idx < len(meta_vals) else ""

        dates, orig_qtys = extract_change_pairs(row, pair_cols)
        corr_qtys, cnt = apply_retroactive_correction(orig_qtys)
        total_corrections += cnt
        max_pair_count = max(max_pair_count, len(dates))

        first_qty = next((v for v in corr_qtys if v is not None), None)
        last_qty = next((v for v in reversed(corr_qtys) if v is not None), None)
        has_change = (
            first_qty is not None
            and last_qty is not None
            and abs(last_qty - first_qty) > 1e-9
        )
        delta = (last_qty - first_qty) if has_change else 0.0

        items.append(
            {
                "row_index": row_i,
                "meta": meta_vals,
                "mgmt_no": mgmt_no,
                "korean_name": korean_name,
                "label": item_label(mgmt_no, korean_name),
                "dates": dates,
                "original": orig_qtys,
                "corrected": corr_qtys,
                "has_change": has_change,
                "delta": delta,
                "first_qty": first_qty,
                "last_qty": last_qty,
                "correction_cells": cnt,
            }
        )

    return {
        "file_path": file_path,
        "file_name": Path(file_path).name,
        "meta_cols": [str(c) for c in meta_cols],
        "pair_cols": [str(c) for c in pair_cols],
        "mgmt_idx": mgmt_idx,
        "name_idx": name_idx,
        "items": items,
        "correction_count": total_corrections,
        "row_count": len(items),
        "max_pair_count": max_pair_count,
    }


def build_analysis_summary(data: dict[str, Any]) -> str:
    """수량 변화가 있는 핵심 품목만 Gemini에 전달할 요약 텍스트 생성."""
    changed_items = [it for it in data["items"] if it["has_change"]]

    lines = [
        f"파일명: {data['file_name']}",
        f"전체 품목 수: {data['row_count']}",
        f"수량 변화 품목 수(분석 대상): {len(changed_items)}",
        f"소급 보정 적용 셀 수: {data['correction_count']}",
        "",
        "※ 최초 대비 최종 재고 수량 변화가 0이거나 데이터가 없는 품목은 제외했습니다.",
        "",
        "=== 핵심 재고 변동 품목 (소급 보정 후) ===",
    ]

    if not changed_items:
        lines.append("(분석 대상 품목 없음)")
        return "\n".join(lines)

    # 절대 변화량 큰 순으로 정렬 후 상위 N건
    ranked = sorted(changed_items, key=lambda x: abs(x["delta"]), reverse=True)
    for it in ranked[:40]:
        timeline = ", ".join(
            f"{d}:{q:g}" if q is not None else f"{d}:-"
            for d, q in zip(it["dates"], it["corrected"])
        )
        lines.append(
            f"- {it['label']}: 최초={it['first_qty']:g} → 최종={it['last_qty']:g} "
            f"(변화={it['delta']:+g}, 보정셀={it['correction_cells']})"
        )
        if timeline:
            lines.append(f"  추이: {timeline}")

    return "\n".join(lines)


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
            result = process_excel(self.file_path)
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(str(e))


class GeminiWorker(QThread):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, api_key: str, summary: str):
        super().__init__()
        self.api_key = api_key
        self.summary = summary

    def run(self) -> None:
        try:
            prompt = f"""당신은 재고 관리 전문가입니다. 아래는 엑셀 재고 데이터에 소급 보정(재고 증가 시점의 차액을 과거 시점에 반영)을 적용한 뒤,
최초 대비 최종 수량이 실제로 변동한 핵심 품목만 추출한 요약입니다.

{self.summary}

위 데이터를 바탕으로 다음 항목을 포함한 **한국어** 종합 분석 리포트를 작성해 주세요:

1. **재고 변동 특징**: 전반적인 재고 추세, 증감 패턴, 품목별 특이사항
2. **이상 징후**: 급격한 변동, 불일치 가능성, 주의가 필요한 품목/시점
3. **적정 재고 관리 가이드**: 발주·안전재고·재고 점검 주기 등 실무 권고사항

마크다운 형식으로 명확하고 구체적으로 작성해 주세요.
변화가 없는 품목은 이미 제외되었으므로, 전달된 핵심 변동 품목 중심으로 분석해 주세요."""

            text = generate_gemini_report(self.api_key, prompt)
            self.finished.emit(text)
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
    """편집형 콤보에서도 입력란 클릭 시 목록이 열리도록 처리."""

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
        layout.setSpacing(4)

        title = QLabel("엑셀 파일을 여기에 드래그 앤 드롭")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("color: #1e3a5f; font-size: 14px; font-weight: 700; border: none; background: transparent;")

        subtitle = QLabel(".xlsx / .xls 지원 · 또는 오른쪽 [파일 선택] 버튼 사용")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setStyleSheet("color: #5b6b7c; font-size: 12px; border: none; background: transparent;")

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
        url = event.mimeData().urls()[0]
        path = url.toLocalFile()
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
        self._show_placeholder("품목을 선택하면 재고 추이 차트가 표시됩니다.")

    def _show_placeholder(self, message: str) -> None:
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

    def plot_item(
        self,
        label: str,
        dates: list[str],
        corrected: list[float | None],
        original: list[float | None] | None = None,
    ) -> None:
        if not dates:
            self._show_placeholder(f"{label}\n표시할 변동기록이 없습니다.")
            return

        y_corr = [v if v is not None else float("nan") for v in corrected]
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        ax.set_facecolor("#fafbfc")

        x = range(len(dates))
        if original is not None:
            y_orig = [v if v is not None else float("nan") for v in original]
            ax.plot(
                x, y_orig, "o--", label="원본 재고",
                color="#c0392b", linewidth=1.6, markersize=5, alpha=0.75,
            )
        ax.plot(
            x, y_corr, "s-", label="소급 보정 후",
            color="#1e3a5f", linewidth=2.2, markersize=6,
        )

        ax.set_xticks(list(x))
        ax.set_xticklabels(dates, rotation=35, ha="right", fontsize=8)
        ax.set_xlabel("변경일자")
        ax.set_ylabel("재고량")
        ax.set_title(f"재고 추이 — {label}", fontsize=12, fontweight="bold", color="#0b1f3a", pad=12)
        ax.legend(loc="best", frameon=False)
        ax.grid(True, axis="y", alpha=0.28, linestyle="--")
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        self.figure.tight_layout()
        self.draw()


# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("재고 분석 및 소급 보정")
        self.setMinimumSize(1140, 760)
        self.resize(1320, 860)

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

        title = QLabel("재고 분석 및 소급 보정")
        title.setObjectName("titleLabel")
        root.addWidget(title)

        # --- Top card: dropzone + API key ---
        top_card = QFrame()
        top_card.setObjectName("card")
        top_layout = QVBoxLayout(top_card)
        top_layout.setContentsMargins(14, 12, 14, 12)
        top_layout.setSpacing(10)

        file_label = QLabel("엑셀 업로드")
        file_label.setObjectName("sectionLabel")
        top_layout.addWidget(file_label)

        file_row = QHBoxLayout()
        file_row.setSpacing(10)
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
        api_row.setSpacing(8)
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

        # --- Tabs ---
        self.tabs = QTabWidget()

        # Tab 1: table
        table_wrap = QWidget()
        table_layout = QVBoxLayout(table_wrap)
        table_layout.setContentsMargins(8, 8, 8, 8)
        hint = QLabel("행을 더블클릭하면 해당 품목의 재고 추이 차트로 이동합니다.")
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

        # Tab 2: chart
        chart_widget = QWidget()
        chart_layout = QVBoxLayout(chart_widget)
        chart_layout.setContentsMargins(10, 10, 10, 10)
        chart_layout.setSpacing(8)

        combo_row = QHBoxLayout()
        combo_label = QLabel("품목 선택")
        combo_label.setObjectName("sectionLabel")
        combo_row.addWidget(combo_label)

        self.item_combo = QComboBox()
        self.item_combo.setEditable(True)
        self.item_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.item_combo.setMinimumWidth(360)
        self.item_combo.setMaxVisibleItems(20)
        self.item_combo.setSizeAdjustPolicy(
            QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon
        )
        line_edit = self.item_combo.lineEdit()
        if line_edit is not None:
            line_edit.setPlaceholderText("관리번호(한글명) 검색/선택 · 클릭하여 목록 열기")
            line_edit.setClearButtonEnabled(True)
            self._combo_popup_filter = ComboPopupFilter(self.item_combo)
            line_edit.installEventFilter(self._combo_popup_filter)

        completer = QCompleter(self.item_combo.model(), self.item_combo)
        completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        completer.setFilterMode(Qt.MatchFlag.MatchContains)
        completer.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
        self.item_combo.setCompleter(completer)

        self.item_combo.currentIndexChanged.connect(self._on_item_combo_changed)
        combo_row.addWidget(self.item_combo, stretch=1)

        btn_open_list = QPushButton("목록")
        btn_open_list.setObjectName("secondaryBtn")
        btn_open_list.setFixedWidth(72)
        btn_open_list.setToolTip("품목 목록 열기")
        btn_open_list.clicked.connect(self.item_combo.showPopup)
        combo_row.addWidget(btn_open_list)
        chart_layout.addLayout(combo_row)

        self.chart = InventoryChart()
        chart_layout.addWidget(self.chart, stretch=1)
        self.tabs.addTab(chart_widget, "재고 추이 차트")

        # Tab 3: AI report
        report_wrap = QWidget()
        report_layout = QVBoxLayout(report_wrap)
        report_layout.setContentsMargins(8, 8, 8, 8)
        self.report_edit = QTextEdit()
        self.report_edit.setReadOnly(True)
        self.report_edit.setPlaceholderText(
            "엑셀 파일을 업로드하고 API 키를 설정하면,\n"
            "수량 변화가 있는 핵심 품목만 추려 AI 분석 리포트를 생성합니다."
        )
        self.report_edit.setFont(QFont("Malgun Gothic", 10))
        report_layout.addWidget(self.report_edit)
        self.tabs.addTab(report_wrap, "AI 분석 리포트")

        root.addWidget(self.tabs, stretch=1)

        # --- Status bar ---
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

    @staticmethod
    def _status_divider() -> QFrame:
        line = QFrame()
        line.setObjectName("statusDivider")
        line.setFrameShape(QFrame.Shape.VLine)
        line.setFixedWidth(1)
        return line

    def _load_api_key(self) -> None:
        config = load_config()
        key = config.get("gemini_api_key", "")
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
        prefix = "API: "
        full_text = f"{prefix}{message}"
        display_text = full_text if len(full_text) <= 100 else full_text[:97] + "..."
        self.status_api.setText(display_text)
        self.status_api.setToolTip(message)
        if ok:
            self.status_api.setStyleSheet("color: #15803d; font-weight: 600;")
        else:
            self.status_api.setStyleSheet("color: #b91c1c; font-weight: 600;")

    def _test_api_connection(self) -> None:
        if self.test_worker is not None and self.test_worker.isRunning():
            self.status_api.setText("API: 연결 확인 중... (이미 진행 중)")
            self.status_api.setStyleSheet("color: #64748b;")
            return

        key = self.api_key_input.text().strip()
        if not key:
            msg = "키 없음 — API Key를 입력해 주세요."
            log_gemini("ERROR", msg)
            self._set_api_status(False, msg)
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

        path, _ = QFileDialog.getOpenFileName(
            self, "엑셀 파일 선택", "", "Excel Files (*.xlsx *.xls)"
        )
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
        self._populate_item_combo(data)
        self._run_ai_analysis(data)
        self.tabs.setCurrentIndex(0)

    def _populate_table(self, data: dict[str, Any]) -> None:
        meta_cols = data["meta_cols"]
        max_pairs = data["max_pair_count"]

        headers = list(meta_cols)
        for i in range(1, max_pairs + 1):
            headers.append(f"변경일자{i}")
            headers.append(f"재고량{i}(원본)")
            headers.append(f"재고량{i}(보정)")

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
                    orig = item["original"][i]
                    corr = item["corrected"][i]
                    orig_text = "" if orig is None else f"{orig:g}"
                    corr_text = "" if corr is None else f"{corr:g}"
                else:
                    date_text, orig_text, corr_text = "", "", ""
                    orig = corr = None

                self.table.setItem(row_idx, col, QTableWidgetItem(date_text))
                self.table.setItem(row_idx, col + 1, QTableWidgetItem(orig_text))

                corr_item = QTableWidgetItem(corr_text)
                if orig is not None and corr is not None and abs(orig - corr) > 1e-9:
                    corr_item.setBackground(highlight)
                self.table.setItem(row_idx, col + 2, corr_item)
                col += 3

            # 행에 관리번호 저장 (더블클릭 연동용)
            first_item = self.table.item(row_idx, 0)
            if first_item is not None:
                first_item.setData(Qt.ItemDataRole.UserRole, item["label"])

    def _populate_item_combo(self, data: dict[str, Any]) -> None:
        self._updating_combo = True
        self.item_combo.blockSignals(True)
        self.item_combo.clear()
        self._item_by_label.clear()

        self.item_combo.addItem("품목을 선택하세요")
        for item in data["items"]:
            label = item["label"]
            # 동일 라벨 중복 시 행 번호로 구분
            if label in self._item_by_label:
                label = f"{label} [행{item['row_index'] + 1}]"
                item = {**item, "label": label}
            self._item_by_label[label] = item
            self.item_combo.addItem(label)

        self.item_combo.setCurrentIndex(0)
        self.item_combo.blockSignals(False)
        self._updating_combo = False
        self.chart._show_placeholder("품목을 선택하면 재고 추이 차트가 표시됩니다.")

    def _on_item_combo_changed(self, index: int) -> None:
        if self._updating_combo or index <= 0:
            return
        label = self.item_combo.currentText().strip()
        item = self._item_by_label.get(label)
        if item:
            self.chart.plot_item(item["label"], item["dates"], item["corrected"], item["original"])

    def _select_item_by_label(self, label: str) -> bool:
        if not label:
            return False
        idx = self.item_combo.findText(label)
        if idx < 0:
            # 관리번호만으로 부분 매칭
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
        first_item = self.table.item(row, 0)
        if first_item is not None:
            label = first_item.data(Qt.ItemDataRole.UserRole)

        if not label and 0 <= row < len(self.inventory_data["items"]):
            label = self.inventory_data["items"][row]["label"]

        self.tabs.setCurrentIndex(1)
        if label and self._select_item_by_label(str(label)):
            return

        # 관리번호 컬럼 직접 조회 폴백
        mgmt_idx = self.inventory_data.get("mgmt_idx", 2)
        mgmt_item = self.table.item(row, mgmt_idx)
        if mgmt_item:
            mgmt_no = mgmt_item.text().strip()
            for i in range(1, self.item_combo.count()):
                if self.item_combo.itemText(i).startswith(mgmt_no):
                    self.item_combo.setCurrentIndex(i)
                    return

    def _run_ai_analysis(self, data: dict[str, Any]) -> None:
        key = self.api_key_input.text().strip()
        changed_count = sum(1 for it in data["items"] if it["has_change"])

        if changed_count == 0:
            self.report_edit.setPlainText(
                "분석 대상 없음\n\n"
                "소급 보정 후 최초 대비 최종 수량 변화가 있는 품목이 없어 "
                "AI 분석을 건너뛰었습니다."
            )
            return

        if not key:
            self.report_edit.setPlainText(
                "API Key가 설정되지 않았습니다.\n"
                "상단에서 Gemini API Key를 입력하고 [저장]을 눌러 주세요.\n\n"
                f"(현재 변동 품목 {changed_count}건이 분석 대기 중입니다.)"
            )
            return

        self.report_edit.setPlainText(
            f"AI 분석 리포트 생성 중... (변동 품목 {changed_count}건)\n잠시만 기다려 주세요."
        )

        summary = build_analysis_summary(data)
        self.gemini_worker = GeminiWorker(key, summary)
        self.gemini_worker.finished.connect(self._on_report_ready)
        self.gemini_worker.error.connect(self._on_report_error)
        self.gemini_worker.start()

    def _on_report_ready(self, text: str) -> None:
        self.report_edit.setMarkdown(text)
        self._set_api_status(True, f"연결됨 ({get_active_gemini_model()})")

    def _on_report_error(self, message: str) -> None:
        self.report_edit.setPlainText(f"AI 분석 생성 실패:\n\n{message}")
        self._set_api_status(False, message)


def main() -> None:
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(APP_STYLE)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
