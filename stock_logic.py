"""생약표준품 재고 엑셀 파싱 · 연도별 소급 보정 · AI 프롬프트 구성."""

from __future__ import annotations

import html
import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Optional, Sequence, Union

import pandas as pd

# 표/식별에 쓰는 기본 메타 (등록일자 제외, 잔고→재고)
META_PREFERRED = [
    "순번",
    "표준품구분",
    "관리번호",
    "한글명",
    "영문명",
    "재고",
    "분양여부",
]

BALANCE_ALIASES = ("재고", "잔고")
STD_TYPE_ALIASES = {
    "대조품": "대조생약",
    "표준생약": "표준생약",
    "지표성분": "지표성분",
    "대조생약": "대조생약",
}
BULK_DECREASE_THRESHOLD = 100
RELIABILITY_LOW_MAX = 2
RELIABILITY_MID_MAX = 4
MANUFACTURE_CANDIDATE_LIMIT = 10
LONG_TERM_LOW_YEARS = 5.0
LONG_TERM_LOW_REL_DROP = 0.05  # 5년+ 구간 상대 감소율 5% 미만 → 저분양
PRICE_COL_KEYWORDS = ("가격",)
ACCELERATION_FORMULA_KO = (
    "분양 가속도 = 최근 3년 연평균 분양량(감소구간만) ÷ 과거 연평균 분양량; "
    "급가속=비율≥2, 증가=비율≥1.25"
)
PRIORITY_FORMULA_KO = (
    "제조우선순위점수 f = 0.40×재고위험도 + 0.25×최근분양속도(정규화) "
    "+ 0.20×분양가속도(정규화) + 0.15×데이터신뢰도\n"
    "  · 재고위험도 = 1 − (예상소진년수/15)  (0~1, 소진년수 없으면 0.05)\n"
    "  · 최근분양속도 정규화 = min(1, 연평균분양량/50)\n"
    "  · 분양가속도 정규화 = min(1, (가속도비율−1)/2)  (비율 없으면 0)\n"
    "  · 데이터신뢰도 = A:1.0 / B:0.75 / C:0.5 / D:0.25"
)
PathLike = Union[str, Path]
PathList = Sequence[PathLike]


def _is_empty(value: Any) -> bool:
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def parse_date(value: Any) -> Optional[date]:
    if _is_empty(value):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, pd.Timestamp):
        if pd.isna(value):
            return None
        return value.date()
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    ts = pd.to_datetime(text, errors="coerce")
    if pd.isna(ts):
        return None
    return ts.date()


def parse_qty(value: Any) -> Optional[float]:
    if _is_empty(value):
        return None
    if isinstance(value, str):
        value = value.strip().replace(",", "")
        if not value:
            return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def format_date(d: date) -> str:
    """호환용 전체 일자 표기 (내부·대량출고 탐지 등)."""
    return d.strftime("%Y-%m-%d")


def format_year(d: date | int) -> str:
    """UI/차트용 연도(YYYY) 표기."""
    if isinstance(d, int):
        return str(d)
    return str(d.year)


def _cell_str(value: Any) -> str:
    if _is_empty(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalize_std_type(raw: Any) -> str:
    """생약(표준생약) → 표준생약, 생약(대조품) → 대조생약 등으로 정제."""
    text = _cell_str(raw)
    if not text:
        return ""
    match = re.search(r"\(([^)]+)\)", text)
    core = match.group(1).strip() if match else text
    core = re.sub(r"^생약\s*", "", core).strip() or core
    return STD_TYPE_ALIASES.get(core, core)


def display_column_name(name: str) -> str:
    """엑셀 원본 컬럼명을 UI 표기명으로 변환 (잔고→재고, 등록일자 제외는 상위에서 처리)."""
    text = str(name).strip()
    if text == "잔고":
        return "재고"
    return text


def discover_timeseries_pairs(columns: list[str]) -> list[tuple[str, str]]:
    """컬럼명에 '변경일자'/'재고량' 키워드가 있는 열을 동적으로 짝지음.

    고정 열 인덱스를 사용하지 않으며, 규격·단위·가격 등 중간 컬럼은 무시한다.
    """
    date_map: dict[int, str] = {}
    qty_map: dict[int, str] = {}

    for col in columns:
        name = str(col).strip()
        if "등록일자" in name:
            continue
        if "변경일자" in name:
            m = re.search(r"변경일자\s*(\d+)", name)
            key = int(m.group(1)) if m else (10_000 + len(date_map))
            date_map[key] = col
        elif "재고량" in name:
            m = re.search(r"재고량\s*(\d+)", name)
            key = int(m.group(1)) if m else (10_000 + len(qty_map))
            qty_map[key] = col

    pairs: list[tuple[str, str]] = []
    for key in sorted(set(date_map) & set(qty_map)):
        pairs.append((date_map[key], qty_map[key]))
    return pairs


def classify_columns(columns: list[str]) -> tuple[list[str], list[tuple[str, str]]]:
    """메타 컬럼과 시계열 쌍을 분리. 등록일자는 메타에서 제외."""
    pairs = discover_timeseries_pairs(columns)
    pair_cols = {c for pair in pairs for c in pair}
    meta_cols: list[str] = []
    for col in columns:
        name = str(col).strip()
        if col in pair_cols:
            continue
        if "등록일자" in name:
            continue
        meta_cols.append(col)

    # 선호 순서 정렬 (잔고/재고 통합)
    ordered: list[str] = []
    used: set[str] = set()
    balance_col = next((c for c in meta_cols if str(c).strip() in BALANCE_ALIASES), None)

    for preferred in META_PREFERRED:
        if preferred == "재고":
            if balance_col and balance_col not in used:
                ordered.append(balance_col)
                used.add(balance_col)
            continue
        hit = next((c for c in meta_cols if str(c).strip() == preferred), None)
        if hit and hit not in used:
            ordered.append(hit)
            used.add(hit)

    for col in meta_cols:
        if col not in used:
            ordered.append(col)
            used.add(col)

    return ordered, pairs


@dataclass
class StockPoint:
    change_date: date
    quantity: float


@dataclass
class StockItem:
    seq: Any = None
    std_type: str = ""
    std_type_raw: Any = None
    manage_no: str = ""
    name_ko: str = ""
    name_en: Any = None
    balance: Any = None
    registered_date: Optional[date] = None
    distributed: Any = None
    unit_price: Optional[float] = None
    extra_meta: dict[str, Any] = field(default_factory=dict)
    raw_points: list[StockPoint] = field(default_factory=list)
    year_end_points: list[StockPoint] = field(default_factory=list)
    corrected_points: list[StockPoint] = field(default_factory=list)
    correction_count: int = 0

    @property
    def label(self) -> str:
        if self.manage_no and self.name_ko:
            return f"{self.manage_no}({self.name_ko})"
        return self.manage_no or self.name_ko or "(미식별)"

    @property
    def first_qty(self) -> Optional[float]:
        if not self.corrected_points:
            return None
        return self.corrected_points[0].quantity

    @property
    def last_qty(self) -> Optional[float]:
        if not self.corrected_points:
            return None
        return self.corrected_points[-1].quantity

    @property
    def current_qty(self) -> Optional[float]:
        """현 재고: 보정 최종값 우선, 없으면 잔고/재고 컬럼."""
        if self.last_qty is not None:
            return self.last_qty
        return parse_qty(self.balance)

    @property
    def stock_value(self) -> Optional[float]:
        """현재 재고 × 가격(원)."""
        qty = self.current_qty
        if qty is None or self.unit_price is None:
            return None
        return float(qty) * float(self.unit_price)

    @property
    def qty_delta(self) -> Optional[float]:
        if self.first_qty is None or self.last_qty is None:
            return None
        return self.last_qty - self.first_qty

    @property
    def has_stock_change(self) -> bool:
        delta = self.qty_delta
        if delta is None:
            return False
        return abs(delta) > 1e-9


def collapse_same_dates(points: list[StockPoint]) -> list[StockPoint]:
    """동일 변경일자에 여러 값이 있으면 마지막 값만 유지 후 날짜 정렬."""
    by_date: dict[date, float] = {}
    for point in points:
        by_date[point.change_date] = point.quantity
    return [StockPoint(d, by_date[d]) for d in sorted(by_date.keys())]


def collapse_to_year_end(points: list[StockPoint]) -> list[StockPoint]:
    """연도별 최종 재고량(해당 연도 마지막 기록일·수량)만 남김."""
    same_day = collapse_same_dates(points)
    year_end: dict[int, StockPoint] = {}
    for point in same_day:
        year_end[point.change_date.year] = point
    return [year_end[y] for y in sorted(year_end.keys())]


def fill_continuous_years(
    points: list[StockPoint],
    *,
    fill_forward: bool = True,
    year_min: int | None = None,
    year_max: int | None = None,
) -> list[StockPoint]:
    """연도 축이 끊기지 않도록 빈 연도를 직전 수량으로 채운다.

    분석(분양속도 등)에는 원본 연도말 포인트를 쓰고, 표/차트 표시에만 사용한다.
    """
    if not points and (year_min is None or year_max is None):
        return []
    by_year = {p.change_date.year: p.quantity for p in points}
    if not by_year and year_min is not None and year_max is not None:
        return []
    y0 = year_min if year_min is not None else min(by_year)
    y1 = year_max if year_max is not None else max(by_year)
    if y1 < y0:
        return []
    result: list[StockPoint] = []
    last: float | None = None
    for y in range(y0, y1 + 1):
        if y in by_year:
            last = by_year[y]
            result.append(StockPoint(date(y, 12, 31), last))
        elif fill_forward and last is not None:
            result.append(StockPoint(date(y, 12, 31), last))
        else:
            result.append(StockPoint(date(y, 12, 31), float("nan")))
    return result


def align_display_series(
    corrected: list[StockPoint],
    original: list[StockPoint],
) -> tuple[list[str], list[float | None], list[float | None]]:
    """보정·원본을 공통 연속 연도축으로 맞춘다. 반환: years, corr_qtys, orig_qtys."""
    years_set = {p.change_date.year for p in corrected} | {
        p.change_date.year for p in original
    }
    if not years_set:
        return [], [], []
    y0, y1 = min(years_set), max(years_set)
    corr_filled = fill_continuous_years(corrected, year_min=y0, year_max=y1)
    orig_map = {p.change_date.year: p.quantity for p in original}
    years = [format_year(p.change_date) for p in corr_filled]
    corr_qtys: list[float | None] = []
    orig_qtys: list[float | None] = []
    last_orig: float | None = None
    for p in corr_filled:
        y = p.change_date.year
        cq = p.quantity
        corr_qtys.append(None if cq != cq else cq)  # NaN → None
        if y in orig_map:
            last_orig = orig_map[y]
            orig_qtys.append(last_orig)
        else:
            orig_qtys.append(last_orig)
    return years, corr_qtys, orig_qtys


def apply_retroactive_correction(points: list[StockPoint]) -> tuple[list[StockPoint], int]:
    """연도별 최종 재고량 기준 소급 보정.

    1) 동일일자 → 마지막 값
    2) 연도별 최종 기록만 추출
    3) 재고가 증가한 연도의 차액만큼 과거 연도 수량을 상향 조정
    """
    year_points = collapse_to_year_end(points)
    if not year_points:
        return [], 0

    quantities = [p.quantity for p in year_points]
    corrected = list(quantities)
    correction_count = 0

    for j in range(1, len(corrected)):
        prev_val = corrected[j - 1]
        curr_val = corrected[j]
        if curr_val > prev_val:
            delta = curr_val - prev_val
            for k in range(j):
                corrected[k] += delta
                correction_count += 1

    result = [
        StockPoint(change_date=p.change_date, quantity=corrected[i])
        for i, p in enumerate(year_points)
    ]
    return result, correction_count


def extract_change_pairs(row: pd.Series, pairs: list[tuple[str, str]]) -> list[StockPoint]:
    """동적 탐색된 변경일자/재고량 쌍에서 유효 포인트만 추출."""
    points: list[StockPoint] = []
    for date_col, qty_col in pairs:
        d = parse_date(row.get(date_col))
        q = parse_qty(row.get(qty_col))
        if d is not None and q is not None:
            points.append(StockPoint(change_date=d, quantity=q))
    return points


def find_balance_value(row: pd.Series, columns: list[str]) -> Any:
    for alias in BALANCE_ALIASES:
        if alias in columns:
            return row.get(alias)
    return None


def find_unit_price(row: pd.Series, columns: list[str]) -> Optional[float]:
    """가격(원) 등 가격 키워드 컬럼에서 단가 추출."""
    for col in columns:
        name = str(col).strip()
        if any(k in name for k in PRICE_COL_KEYWORDS):
            q = parse_qty(row.get(col))
            if q is not None and q >= 0:
                return q
    return None


def _extract_unit_price_from_meta(extra_meta: dict[str, Any]) -> Optional[float]:
    for key, val in extra_meta.items():
        if any(k in str(key) for k in PRICE_COL_KEYWORDS):
            q = parse_qty(val)
            if q is not None and q >= 0:
                return q
    return None


def build_corrected_dataframe(
    items: list[StockItem],
    meta_cols: list[str],
) -> pd.DataFrame:
    """소급 보정 최종 수치만 반영한 표용 DataFrame (연도 YYYY · 연속 연도축)."""
    display_series = [align_display_series(it.corrected_points, it.year_end_points) for it in items]
    max_points = max((len(s[0]) for s in display_series), default=0)
    display_meta = [display_column_name(c) for c in meta_cols]
    records: list[dict[str, Any]] = []

    for it, (years, corr_qtys, _orig_qtys) in zip(items, display_series):
        rec: dict[str, Any] = {}
        for src, disp in zip(meta_cols, display_meta):
            key = str(src).strip()
            if key in BALANCE_ALIASES:
                rec[disp] = it.balance
            elif key == "표준품구분":
                rec[disp] = it.std_type
            elif key == "관리번호":
                rec[disp] = it.manage_no
            elif key == "한글명":
                rec[disp] = it.name_ko
            elif key == "영문명":
                rec[disp] = it.name_en
            elif key == "순번":
                rec[disp] = it.seq
            elif key == "분양여부":
                rec[disp] = it.distributed
            else:
                rec[disp] = it.extra_meta.get(src, "")

        for idx in range(max_points):
            if idx < len(years):
                rec[f"변경일자{idx + 1}"] = years[idx]
                rec[f"재고량{idx + 1}"] = corr_qtys[idx]
            else:
                rec[f"변경일자{idx + 1}"] = None
                rec[f"재고량{idx + 1}"] = None
        records.append(rec)

    columns = list(display_meta)
    for idx in range(max_points):
        columns += [f"변경일자{idx + 1}", f"재고량{idx + 1}"]
    return pd.DataFrame(records, columns=columns)


def _item_merge_key(item: StockItem) -> str:
    if item.manage_no:
        return f"no:{item.manage_no}"
    return f"name:{item.name_ko}"


def merge_stock_items(groups: Sequence[Sequence[StockItem]]) -> list[StockItem]:
    """여러 파일의 품목을 관리번호(없으면 한글명) 기준으로 통합."""
    merged: dict[str, StockItem] = {}
    order: list[str] = []

    for group in groups:
        for it in group:
            key = _item_merge_key(it)
            if key not in merged:
                merged[key] = StockItem(
                    seq=it.seq,
                    std_type=it.std_type,
                    std_type_raw=it.std_type_raw,
                    manage_no=it.manage_no,
                    name_ko=it.name_ko,
                    name_en=it.name_en,
                    balance=it.balance,
                    registered_date=it.registered_date,
                    distributed=it.distributed,
                    unit_price=it.unit_price,
                    extra_meta=dict(it.extra_meta),
                    raw_points=list(it.raw_points),
                    year_end_points=list(it.year_end_points),
                    corrected_points=list(it.corrected_points),
                    correction_count=it.correction_count,
                )
                order.append(key)
                continue

            existing = merged[key]
            existing.raw_points = list(existing.raw_points) + list(it.raw_points)
            existing.year_end_points = collapse_to_year_end(existing.raw_points)
            existing.corrected_points, existing.correction_count = apply_retroactive_correction(
                existing.raw_points
            )
            if it.std_type:
                existing.std_type = it.std_type
                existing.std_type_raw = it.std_type_raw or existing.std_type_raw
            if it.name_ko:
                existing.name_ko = it.name_ko
            if not _is_empty(it.name_en):
                existing.name_en = it.name_en
            if not _is_empty(it.balance):
                existing.balance = it.balance
            if it.unit_price is not None:
                existing.unit_price = it.unit_price
            if not _is_empty(it.distributed):
                existing.distributed = it.distributed
            if it.registered_date and (
                existing.registered_date is None or it.registered_date > existing.registered_date
            ):
                existing.registered_date = it.registered_date
            existing.extra_meta.update({k: v for k, v in it.extra_meta.items() if not _is_empty(v)})
            if existing.seq is None:
                existing.seq = it.seq

    return [merged[k] for k in order]


def _unify_meta_cols(items: list[StockItem]) -> list[str]:
    """통합 품목 목록에서 표시용 메타 컬럼 순서 결정."""
    present = {"관리번호", "한글명"}
    for it in items:
        if it.seq is not None:
            present.add("순번")
        if it.std_type:
            present.add("표준품구분")
        if not _is_empty(it.name_en):
            present.add("영문명")
        if not _is_empty(it.balance):
            present.add("재고")
        if not _is_empty(it.distributed):
            present.add("분양여부")
        for k in it.extra_meta:
            present.add(str(k))
    ordered: list[str] = []
    for pref in META_PREFERRED:
        if pref in present:
            ordered.append(pref)
            present.discard(pref)
    ordered.extend(sorted(present - set(ordered)))
    return ordered


def load_stock_excel(path: str) -> tuple[pd.DataFrame, list[StockItem]]:
    """엑셀을 읽어 연도별 소급 보정된 StockItem 목록과 표용 DataFrame을 반환."""
    items = load_stock_items(path)
    meta_cols = _unify_meta_cols(items)
    return build_corrected_dataframe(items, meta_cols), items


def load_stock_items(path: PathLike) -> list[StockItem]:
    """단일 엑셀에서 StockItem 목록만 로드."""
    path = str(path)
    df = pd.read_excel(path, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in ("관리번호", "한글명") if c not in df.columns]
    if missing:
        raise ValueError(f"필수 컬럼이 없습니다 ({Path(path).name}): {', '.join(missing)}")

    meta_cols, pairs = classify_columns(list(df.columns))
    if not pairs:
        raise ValueError(
            f"'{Path(path).name}'에서 '변경일자'/'재고량' 시계열 컬럼 쌍을 찾을 수 없습니다."
        )

    known = {"순번", "표준품구분", "관리번호", "한글명", "영문명", "재고", "잔고", "분양여부"}
    reg_col = next((c for c in df.columns if "등록일자" in str(c)), None)
    items: list[StockItem] = []

    for _, row in df.iterrows():
        manage_no = _cell_str(row.get("관리번호"))
        name_ko = _cell_str(row.get("한글명"))
        if not manage_no and not name_ko:
            continue

        raw_points = extract_change_pairs(row, pairs)
        year_end = collapse_to_year_end(raw_points)
        corrected, corr_cnt = apply_retroactive_correction(raw_points)
        registered = parse_date(row.get(reg_col)) if reg_col else None

        extra = {
            c: row.get(c)
            for c in meta_cols
            if str(c).strip() not in known
        }

        items.append(
            StockItem(
                seq=row.get("순번"),
                std_type=normalize_std_type(row.get("표준품구분")),
                std_type_raw=row.get("표준품구분"),
                manage_no=manage_no,
                name_ko=name_ko,
                name_en=row.get("영문명"),
                balance=find_balance_value(row, list(df.columns)),
                registered_date=registered,
                distributed=row.get("분양여부"),
                unit_price=find_unit_price(row, list(df.columns)),
                extra_meta=extra,
                raw_points=raw_points,
                year_end_points=year_end,
                corrected_points=corrected,
                correction_count=corr_cnt,
            )
        )

    return items


def process_excel(file_path: PathLike) -> dict[str, Any]:
    """단일 파일 GUI 처리 (다중 파일 API의 래퍼)."""
    return process_excels([file_path])


def process_excels(file_paths: PathList) -> dict[str, Any]:
    """하나 이상의 엑셀을 통합 데이터셋으로 결합해 GUI용 결과를 반환."""
    paths = [str(p) for p in file_paths]
    if not paths:
        raise ValueError("업로드할 엑셀 파일이 없습니다.")

    groups = [load_stock_items(p) for p in paths]
    items = merge_stock_items(groups) if len(groups) > 1 else list(groups[0])
    meta_cols = _unify_meta_cols(items)
    table_df = build_corrected_dataframe(items, meta_cols)

    display_series = [
        align_display_series(it.corrected_points, it.year_end_points) for it in items
    ]
    max_pairs = max((len(s[0]) for s in display_series), default=0)
    total_corrections = sum(it.correction_count for it in items)
    categories = sorted({it.std_type for it in items if it.std_type})

    ui_items: list[dict[str, Any]] = []
    for idx, (it, (years, corr_qtys, orig_qtys)) in enumerate(zip(items, display_series)):
        meta_vals = []
        for col in meta_cols:
            if col == "재고":
                meta_vals.append(_cell_str(it.balance))
            elif col == "표준품구분":
                meta_vals.append(it.std_type)
            elif col == "관리번호":
                meta_vals.append(it.manage_no)
            elif col == "한글명":
                meta_vals.append(it.name_ko)
            elif col == "영문명":
                meta_vals.append(_cell_str(it.name_en))
            elif col == "순번":
                meta_vals.append(_cell_str(it.seq))
            elif col == "분양여부":
                meta_vals.append(_cell_str(it.distributed))
            else:
                meta_vals.append(_cell_str(it.extra_meta.get(col, "")))

        ui_items.append(
            {
                "row_index": idx,
                "meta": meta_vals,
                "mgmt_no": it.manage_no,
                "korean_name": it.name_ko,
                "std_type": it.std_type,
                "label": it.label,
                "dates": years,
                "original_dates": years,
                "original": orig_qtys,
                "corrected": corr_qtys,
                "has_change": it.has_stock_change,
                "delta": it.qty_delta if it.qty_delta is not None else 0.0,
                "first_qty": it.first_qty,
                "last_qty": it.last_qty,
                "correction_cells": it.correction_count,
                "stock_item": it,
            }
        )

    names = [Path(p).name for p in paths]
    return {
        "file_path": paths[0],
        "file_paths": paths,
        "file_name": names[0] if len(names) == 1 else f"{len(names)}개 파일",
        "file_names": names,
        "meta_cols": meta_cols,
        "mgmt_idx": meta_cols.index("관리번호") if "관리번호" in meta_cols else 0,
        "items": ui_items,
        "stock_items": items,
        "table_df": table_df,
        "correction_count": total_corrections,
        "row_count": len(items),
        "max_pair_count": max_pairs,
        "categories": categories,
    }


def items_for_ai_analysis(items: list[StockItem]) -> list[StockItem]:
    return [it for it in items if it.has_stock_change]


def detect_bulk_decrease_dates(items: list[StockItem], threshold: int = BULK_DECREASE_THRESHOLD) -> list[str]:
    """동일 날짜에 threshold개 이상 품목이 감소한 날짜(연구과제 대량출고) 탐지."""
    decrease_counter: Counter[str] = Counter()
    for it in items:
        pts = it.corrected_points
        for i in range(1, len(pts)):
            if pts[i].quantity < pts[i - 1].quantity - 1e-9:
                decrease_counter[format_date(pts[i].change_date)] += 1
    return sorted(d for d, n in decrease_counter.items() if n >= threshold)


def decrease_only_rate_stats(points: list[StockPoint]) -> dict[str, Any]:
    """증가 구간(전수조사/추가제조/반납 등)을 제외한 감소 구간만으로 분양속도 산출."""
    total_drop = 0.0
    total_days = 0
    increase_segments = 0
    decrease_segments = 0
    for i in range(1, len(points)):
        delta = points[i - 1].quantity - points[i].quantity
        days = (points[i].change_date - points[i - 1].change_date).days
        if days <= 0:
            continue
        if delta > 1e-9:
            total_drop += delta
            total_days += days
            decrease_segments += 1
        elif delta < -1e-9:
            increase_segments += 1
    if total_days <= 0 or total_drop <= 1e-9:
        return {
            "annual_rate": None,
            "total_drop": total_drop,
            "decrease_years": 0.0,
            "increase_segments": increase_segments,
            "decrease_segments": decrease_segments,
        }
    years = total_days / 365.25
    return {
        "annual_rate": total_drop / years,
        "total_drop": total_drop,
        "decrease_years": years,
        "increase_segments": increase_segments,
        "decrease_segments": decrease_segments,
    }


def _recent_vs_past_decrease_rates(
    points: list[StockPoint],
    recent_years: float = 3.0,
) -> tuple[Optional[float], Optional[float]]:
    """최근 N년 vs 과거 감소 구간 연평균 속도.

    Returns:
        (past_annual_rate, recent_annual_rate)
        cutoff = last_date - recent_years 기준으로 구간을 나눔.
    """
    if len(points) < 2:
        return None, None
    last_date = points[-1].change_date
    cutoff = last_date - timedelta(days=int(round(recent_years * 365.25)))

    split_idx = None
    for i, p in enumerate(points):
        if p.change_date >= cutoff:
            split_idx = i
            break

    if split_idx is None:
        # 전 구간이 과거(최근 구간 없음)
        past = decrease_only_rate_stats(points)["annual_rate"]
        return past, None
    if split_idx == 0:
        # 전 구간이 최근(과거 구간 없음) — 경계 직전 포인트가 없으면 과거 None
        recent = decrease_only_rate_stats(points)["annual_rate"]
        return None, recent

    # 과거: cutoff 이전(+경계점), 최근: 경계 직전 포인트부터(연속 구간 포함)
    past_pts = points[: split_idx + 1]
    recent_pts = points[split_idx - 1 :]
    past = decrease_only_rate_stats(past_pts)["annual_rate"] if len(past_pts) >= 2 else None
    recent = decrease_only_rate_stats(recent_pts)["annual_rate"] if len(recent_pts) >= 2 else None
    return past, recent


def _half_period_decrease_rate(points: list[StockPoint]) -> tuple[Optional[float], Optional[float]]:
    """호환용: 전반/후반 분할. 신규 로직은 `_recent_vs_past_decrease_rates` 사용."""
    if len(points) < 3:
        return None, None
    mid = len(points) // 2
    early = decrease_only_rate_stats(points[: mid + 1])["annual_rate"]
    late = decrease_only_rate_stats(points[mid:])["annual_rate"]
    return early, late


def acceleration_from_rates(
    early_rate: Optional[float],
    late_rate: Optional[float],
) -> dict[str, Any]:
    """분양 가속도: 급가속 / 증가 / 안정 / 감소.

    비율 = 최근(late) / 과거(early). 급가속=비율≥2, 증가=비율≥1.25,
    감소=비율≤0.75, 그 외 안정. is_surge는 급가속만 True.
    """
    if early_rate is None and late_rate is None:
        return {"label": "안정", "ratio": None, "is_surge": False}
    if early_rate is None or early_rate <= 1e-9:
        if late_rate is not None and late_rate > 1e-9:
            return {"label": "급가속", "ratio": 3.0, "is_surge": True}
        return {"label": "안정", "ratio": None, "is_surge": False}
    if late_rate is None or late_rate <= 1e-9:
        return {"label": "감소", "ratio": 0.0, "is_surge": False}

    ratio = float(late_rate) / float(early_rate)
    if ratio >= 2.0:
        label = "급가속"
    elif ratio >= 1.25:
        label = "증가"
    elif ratio <= 0.75:
        label = "감소"
    else:
        label = "안정"
    return {"label": label, "ratio": ratio, "is_surge": label == "급가속"}


def compute_data_reliability(item: StockItem) -> dict[str, Any]:
    """수집 횟수 + 관측 간격 기반 신뢰도 등급 A~D."""
    pts = collapse_same_dates(item.raw_points)
    n = len(pts)
    avg_gap_years: Optional[float] = None
    if n >= 2:
        gaps = [
            (pts[i].change_date - pts[i - 1].change_date).days / 365.25
            for i in range(1, n)
        ]
        avg_gap_years = sum(gaps) / len(gaps) if gaps else None

    if n >= 6:
        grade, score = "A", 1.0
    elif n >= 4:
        grade, score = "B", 0.75
    elif n >= 3:
        grade, score = "C", 0.5
    else:
        grade, score = "D", 0.25

    if avg_gap_years is not None and avg_gap_years > 3.5 and grade in ("A", "B"):
        grade = "B" if grade == "A" else "C"
        score = 0.75 if grade == "B" else 0.5
    if avg_gap_years is not None and avg_gap_years > 5.0 and grade == "C":
        grade, score = "D", 0.25

    labels = {
        "A": "신뢰도 A (높음)",
        "B": "신뢰도 B (양호)",
        "C": "신뢰도 C (보통)",
        "D": "신뢰도 D (낮음·데이터 부족)",
    }
    return {
        "count": n,
        "grade": grade,
        "level": grade,
        "label": labels[grade],
        "score": float(score),
        "avg_gap_years": round(avg_gap_years, 3) if avg_gap_years is not None else None,
    }


def depletion_bucket(years_left: Optional[float]) -> str:
    """소진 예상 기간 카테고리 (최대 15년 기준)."""
    if years_left is None:
        return "15년 이상/안정"
    if years_left <= 1:
        months = max(1, int(round(years_left * 12)))
        return f"1년 이내 ({months}개월)"
    if years_left <= 3:
        return "1~3년 이내"
    if years_left <= 5:
        return "3~5년 이내"
    if years_left <= 10:
        return "5~10년 이내"
    if years_left <= 15:
        return "10~15년 이내"
    return "15년 이상/안정"


def format_deplete_ym(last_date: date, years_left: float) -> str:
    """현재 추세 기준 예상 소진 시점 (YYYY년 M월)."""
    days = max(0, int(round(years_left * 365.25)))
    target = last_date + timedelta(days=days)
    return f"{target.year}년 {target.month}월"


def _priority_components(
    *,
    years_left: Optional[float],
    annual_rate: Optional[float],
    acceleration_ratio: Optional[float],
    reliability_score: float,
) -> dict[str, float]:
    """제조우선순위 구성요소: risk / speed_n / accel_n / rel / score."""
    if years_left is None:
        risk = 0.05
    elif years_left <= 0:
        risk = 1.0
    else:
        risk = max(0.0, min(1.0, 1.0 - (years_left / 15.0)))

    speed_n = 0.0
    if annual_rate is not None and annual_rate > 0:
        speed_n = max(0.0, min(1.0, float(annual_rate) / 50.0))

    accel_n = 0.0
    if acceleration_ratio is not None and acceleration_ratio > 0:
        accel_n = max(0.0, min(1.0, (float(acceleration_ratio) - 1.0) / 2.0))

    rel = max(0.0, min(1.0, float(reliability_score)))
    score = round(0.40 * risk + 0.25 * speed_n + 0.20 * accel_n + 0.15 * rel, 4)
    return {
        "risk": round(risk, 4),
        "speed_n": round(speed_n, 4),
        "accel_n": round(accel_n, 4),
        "rel": round(rel, 4),
        "score": score,
    }


def manufacturing_priority_score(
    *,
    years_left: Optional[float],
    annual_rate: Optional[float],
    acceleration_ratio: Optional[float],
    reliability_score: float,
) -> float:
    """제조우선순위점수 = f(재고위험도, 최근 분양속도, 분양 가속도, 데이터 신뢰도)."""
    return _priority_components(
        years_left=years_left,
        annual_rate=annual_rate,
        acceleration_ratio=acceleration_ratio,
        reliability_score=reliability_score,
    )["score"]


def is_long_term_low_distribution(item: StockItem, annual_rate: Optional[float]) -> bool:
    """최근 5년 이상 감소가 거의 없는 장기 저분양/과다재고."""
    pts = item.corrected_points
    if len(pts) < 2 or item.first_qty is None or item.last_qty is None:
        return False
    span = (pts[-1].change_date - pts[0].change_date).days / 365.25
    if span < LONG_TERM_LOW_YEARS:
        return False
    if abs(item.first_qty) > 1e-12:
        rel_drop = (item.first_qty - item.last_qty) / item.first_qty
        if rel_drop < LONG_TERM_LOW_REL_DROP:
            return True
    if annual_rate is None or annual_rate < 1.0:
        return True
    return False


def estimate_depletion(item: StockItem) -> dict[str, Any]:
    """분양 속도(증가구간 제외)·가속도·신뢰도·우선순위 산출."""
    if item.unit_price is None:
        item.unit_price = _extract_unit_price_from_meta(item.extra_meta)

    reliability = compute_data_reliability(item)
    pts = item.corrected_points
    empty = {
        "speed": "데이터부족",
        "annual_rate": None,
        "years_left": None,
        "deplete_within_2y": False,
        "deplete_within_5y": False,
        "recent_surge": False,
        "rate_change_ratio": None,
        "acceleration": "안정",
        "acceleration_ratio": None,
        "early_rate": None,  # past
        "late_rate": None,  # recent
        "past_rate": None,
        "recent_rate": None,
        "long_term_low": False,
        "increase_segments_excluded": 0,
        "deplete_ym": None,
        "depletion_category": "15년 이상/안정",
        "reliability": reliability,
        "stock_risk": 0.0,
        "priority_score": manufacturing_priority_score(
            years_left=None,
            annual_rate=None,
            acceleration_ratio=None,
            reliability_score=reliability["score"],
        ),
        "unit_price": item.unit_price,
        "stock_value": item.stock_value,
        "as_of_year": date.today().year,
    }
    if len(pts) < 2 or item.first_qty is None or item.last_qty is None:
        return empty

    dec = decrease_only_rate_stats(pts)
    annual_rate = dec["annual_rate"]
    # early/late = past/recent (최근 3년 vs 과거)
    early_rate, late_rate = _recent_vs_past_decrease_rates(pts, recent_years=3.0)
    accel = acceleration_from_rates(early_rate, late_rate)
    rate_change_ratio = accel["ratio"]
    recent_surge = bool(accel["label"] == "급가속")

    if annual_rate is None or annual_rate <= 1e-9:
        speed = "느림"
        years_left = None
        deplete_5 = False
        deplete_2 = False
        deplete_ym = None
    else:
        years_left = item.last_qty / annual_rate if item.last_qty > 0 else 0.0
        if annual_rate >= 40:
            speed = "빠름"
        elif annual_rate >= 10:
            speed = "보통"
        else:
            speed = "느림"
        deplete_5 = years_left is not None and years_left <= 5.0
        deplete_2 = years_left is not None and years_left <= 2.0
        deplete_ym = format_deplete_ym(pts[-1].change_date, float(years_left))

    category = depletion_bucket(years_left)
    long_term_low = is_long_term_low_distribution(item, annual_rate)
    comps = _priority_components(
        years_left=years_left,
        annual_rate=annual_rate,
        acceleration_ratio=accel["ratio"],
        reliability_score=reliability["score"],
    )
    priority = comps["score"]
    stock_risk = comps["risk"]

    return {
        "speed": speed,
        "annual_rate": annual_rate,
        "years_left": years_left,
        "deplete_within_2y": deplete_2,
        "deplete_within_5y": deplete_5,
        "recent_surge": recent_surge,
        "rate_change_ratio": rate_change_ratio,
        "acceleration": accel["label"],
        "acceleration_ratio": accel["ratio"],
        "early_rate": early_rate,
        "late_rate": late_rate,
        "past_rate": early_rate,
        "recent_rate": late_rate,
        "long_term_low": long_term_low,
        "increase_segments_excluded": dec["increase_segments"],
        "deplete_ym": deplete_ym,
        "depletion_category": category,
        "reliability": reliability,
        "stock_risk": stock_risk,
        "speed_n": comps["speed_n"],
        "accel_n": comps["accel_n"],
        "priority_score": priority,
        "unit_price": item.unit_price,
        "stock_value": item.stock_value,
        "as_of_year": date.today().year,
    }


def group_by_depletion_category(items: list[StockItem]) -> dict[str, list[str]]:
    """소진 기간 카테고리별 품목 라벨 목록."""
    order = [
        "1년 이내",
        "1~3년 이내",
        "3~5년 이내",
        "5~10년 이내",
        "10~15년 이내",
        "15년 이상/안정",
    ]
    buckets: dict[str, list[str]] = {k: [] for k in order}
    for it in items_for_ai_analysis(items):
        stats = estimate_depletion(it)
        cat = stats["depletion_category"]
        key = cat
        if cat.startswith("1년 이내"):
            key = "1년 이내"
        if key not in buckets:
            buckets[key] = []
        buckets[key].append(f"{it.label} [{cat}]")
    return buckets


def select_manufacture_candidates(
    items: list[StockItem],
    limit_per_type: int = MANUFACTURE_CANDIDATE_LIMIT,
) -> dict[str, list[dict[str, Any]]]:
    """차년도 제조검토대상: 점수 높고 5년 이내 소진 예상 표준생약/지표성분."""
    result: dict[str, list[dict[str, Any]]] = {"표준생약": [], "지표성분": []}
    scored: list[tuple[float, StockItem, dict[str, Any]]] = []
    for it in items:
        if it.std_type not in result:
            continue
        stats = estimate_depletion(it)
        if not stats["deplete_within_5y"]:
            continue
        scored.append((float(stats["priority_score"]), it, stats))

    scored.sort(key=lambda x: x[0], reverse=True)
    for score, it, stats in scored:
        bucket = result[it.std_type]
        if len(bucket) >= limit_per_type:
            continue
        rel = stats["reliability"]
        rel_score = float(rel.get("score", 0.25)) if isinstance(rel, dict) else 0.25
        comps = _priority_components(
            years_left=stats["years_left"],
            annual_rate=stats["annual_rate"],
            acceleration_ratio=stats["acceleration_ratio"],
            reliability_score=rel_score,
        )
        bucket.append(
            {
                "label": it.label,
                "manage_no": it.manage_no,
                "name_ko": it.name_ko,
                "std_type": it.std_type,
                "priority_score": score,
                "stock_risk": comps["risk"],
                "speed_n": comps["speed_n"],
                "accel_n": comps["accel_n"],
                "reliability_score": comps["rel"],
                "reliability_grade": rel.get("grade") if isinstance(rel, dict) else None,
                "reliability": rel["label"] if isinstance(rel, dict) else str(rel),
                "early_rate": stats.get("early_rate"),
                "late_rate": stats.get("late_rate"),
                "past_rate": stats.get("past_rate", stats.get("early_rate")),
                "recent_rate": stats.get("recent_rate", stats.get("late_rate")),
                "acceleration_ratio": stats["acceleration_ratio"],
                "annual_rate": stats["annual_rate"],
                "years_left": stats["years_left"],
                "deplete_ym": stats["deplete_ym"],
                "depletion_category": stats["depletion_category"],
                "last_qty": it.last_qty,
                "acceleration": stats["acceleration"],
            }
        )
    return result


def select_monitoring_targets(items: list[StockItem], limit: int = 30) -> list[dict[str, Any]]:
    """분양 가속도 급증(급가속/증가) 모니터링 대상.

    급가속은 전량 반환(limit 미적용). 증가 항목은 급가속 뒤에 이어서 포함
    (limit이 있으면 증가 구간에만 상한 적용, 기본은 사실상 무제한).
    """
    surge_rows: list[dict[str, Any]] = []
    increase_rows: list[dict[str, Any]] = []
    for it in items:
        stats = estimate_depletion(it)
        accel = stats.get("acceleration")
        if accel not in ("급가속", "증가"):
            continue
        row = {
            "label": it.label,
            "manage_no": it.manage_no,
            "std_type": it.std_type,
            "acceleration": stats["acceleration"],
            "acceleration_ratio": stats["acceleration_ratio"],
            "rate_change_ratio": stats["rate_change_ratio"],
            "annual_rate": stats["annual_rate"],
            "years_left": stats["years_left"],
            "deplete_ym": stats["deplete_ym"],
            "reliability": stats["reliability"]["label"],
            "priority_score": stats["priority_score"],
            "indicators": (
                f"가속도={stats['acceleration']}"
                + (
                    f", 최근/과거비={stats['acceleration_ratio']:.2f}"
                    if stats["acceleration_ratio"] is not None
                    else ""
                )
            ),
        }
        if accel == "급가속":
            surge_rows.append(row)
        else:
            increase_rows.append(row)

    def _sort_key(r: dict[str, Any]) -> tuple[float, float]:
        return (float(r["acceleration_ratio"] or 0), float(r["priority_score"] or 0))

    surge_rows.sort(key=_sort_key, reverse=True)
    increase_rows.sort(key=_sort_key, reverse=True)
    # 급가속 전량 + 증가 전량 (limit은 API 호환용, 급가속 하드캡 없음)
    _ = limit
    return surge_rows + increase_rows


def select_long_term_low_items(items: list[StockItem], limit: int = 40) -> list[dict[str, Any]]:
    """장기 저분양/과다재고 — 차기 제조 시 수량 하향 조정 권고."""
    rows: list[dict[str, Any]] = []
    for it in items:
        stats = estimate_depletion(it)
        if not stats.get("long_term_low"):
            continue
        rows.append(
            {
                "label": it.label,
                "manage_no": it.manage_no,
                "std_type": it.std_type,
                "annual_rate": stats["annual_rate"],
                "years_left": stats["years_left"],
                "stock_value": stats.get("stock_value"),
                "reliability": stats["reliability"]["label"],
                "recommendation": "차기 제조 시 수량 하향 조정 검토",
            }
        )
    rows.sort(key=lambda r: float(r["stock_value"] or 0), reverse=True)
    return rows[:limit]


def compute_inventory_valuation(items: list[StockItem]) -> dict[str, Any]:
    """현재 재고 × 가격(원) 환산액 집계."""
    by_type: dict[str, float] = defaultdict(float)
    rows: list[dict[str, Any]] = []
    missing_price = 0
    for it in items:
        if it.unit_price is None:
            it.unit_price = _extract_unit_price_from_meta(it.extra_meta)
        val = it.stock_value
        if val is None:
            missing_price += 1
            continue
        tkey = it.std_type or "미분류"
        by_type[tkey] += val
        rows.append(
            {
                "label": it.label,
                "manage_no": it.manage_no,
                "std_type": tkey,
                "qty": it.current_qty,
                "unit_price": it.unit_price,
                "stock_value": val,
            }
        )
    rows.sort(key=lambda r: float(r["stock_value"]), reverse=True)
    total = sum(by_type.values())
    type_order = ["표준생약", "지표성분", "대조생약", "미분류"]
    by_type_ordered: dict[str, float] = {}
    for k in type_order:
        if k in by_type:
            by_type_ordered[k] = by_type[k]
    for k, v in by_type.items():
        if k not in by_type_ordered:
            by_type_ordered[k] = v
    return {
        "total_value": total,
        "by_type": dict(by_type_ordered),
        "top20": rows[:20],
        "priced_count": len(rows),
        "missing_price_count": missing_price,
    }


def _fmt_money(value: float | None) -> str:
    if value is None:
        return "-"
    return f"{value:,.0f}원"


def build_kpi_dashboard(items: list[StockItem], flags: dict[str, Any] | None = None) -> dict[str, Any]:
    """핵심 KPI + 자동 종합 의견(5줄 내외)."""
    flags = flags or {}
    valuation = flags.get("valuation") or compute_inventory_valuation(items)
    by_code = flags.get("by_code") or {}

    managed = len(items)
    deplete_2y = 0
    # deplete_codes가 없으면 직접 집계 (estimate_depletion과 동일 기준)
    if flags.get("deplete_codes") is not None:
        deplete_5y = len(flags.get("deplete_codes") or [])
    else:
        deplete_5y = 0
    accel_n = 0
    low_n = 0
    grade_ab = 0
    for it in items:
        stats = by_code.get(it.manage_no) or estimate_depletion(it)
        if stats.get("deplete_within_2y"):
            deplete_2y += 1
        if flags.get("deplete_codes") is None and stats.get("deplete_within_5y"):
            deplete_5y += 1
        if stats.get("acceleration") in ("급가속", "증가"):
            accel_n += 1
        if stats.get("long_term_low"):
            low_n += 1
        rel = stats.get("reliability") or {}
        if isinstance(rel, dict) and rel.get("grade") in ("A", "B"):
            grade_ab += 1

    mfg = flags.get("manufacture_candidates")
    if mfg is None:
        mfg = select_manufacture_candidates(items)
    mfg_n = len(mfg.get("표준생약") or []) + len(mfg.get("지표성분") or [])
    total_value = float(valuation.get("total_value") or 0)
    by_type = valuation.get("by_type") or {}

    kpis = [
        {"key": "managed", "label": "대상품목 수", "value": managed, "display": f"{managed}종"},
        {"key": "deplete_2y", "label": "2년 내 소진예상", "value": deplete_2y, "display": f"{deplete_2y}종"},
        {"key": "deplete_5y", "label": "5년 내 소진예상", "value": deplete_5y, "display": f"{deplete_5y}종"},
        {"key": "manufacture", "label": "제조 우선검토 수", "value": mfg_n, "display": f"{mfg_n}종"},
        {"key": "accel", "label": "분양 가속 품목 수", "value": accel_n, "display": f"{accel_n}종"},
        {"key": "low_dist", "label": "장기 저분양 품목 수", "value": low_n, "display": f"{low_n}종"},
        {
            "key": "total_value",
            "label": "총 분양 환산금액",
            "value": total_value,
            "display": _fmt_money(total_value),
        },
        {
            "key": "reliability_ab",
            "label": "신뢰도 A·B 품목",
            "value": grade_ab,
            "display": f"{grade_ab}종",
        },
        {
            "key": "priced",
            "label": "가격 반영 품목",
            "value": valuation.get("priced_count", 0),
            "display": f"{valuation.get('priced_count', 0)}종",
        },
    ]

    summary_lines = [
        f"대상 품목 {managed}종을 기준으로 재고·분양 지표를 산출했습니다.",
        f"2년 내 소진 예상 {deplete_2y}종, 5년 내 소진 예상 {deplete_5y}종이며 제조 우선검토 후보는 {mfg_n}종입니다.",
        f"분양 가속도(급가속·증가) 품목은 {accel_n}종, 장기 저분양/과다재고 후보는 {low_n}종입니다.",
        f"현 재고 기준 분양금액 환산 총액은 {_fmt_money(total_value)}입니다"
        + (
            f" (표준생약 {_fmt_money(by_type.get('표준생약'))}, "
            f"지표성분 {_fmt_money(by_type.get('지표성분'))}, "
            f"대조생약 {_fmt_money(by_type.get('대조생약'))})."
            if by_type
            else "."
        ),
        "아래 수치는 사전 정량 산출 결과이며, 데이터에 없는 원인 단정은 하지 않습니다.",
    ]

    return {
        "kpis": kpis,
        "summary_lines": summary_lines,
        "valuation": valuation,
    }


def format_kpi_dashboard_markdown(dashboard: dict[str, Any]) -> str:
    """리포트 최상단용 KPI 대시보드 마크다운."""
    lines = [
        "## 1페이지 요약 대시보드 (핵심 KPI)",
        "",
        "| 지표 | 값 |",
        "| --- | --- |",
    ]
    for kpi in dashboard.get("kpis") or []:
        lines.append(f"| {kpi['label']} | {kpi['display']} |")
    lines.append("")
    lines.append("### 자동 종합 의견")
    for s in dashboard.get("summary_lines") or []:
        lines.append(f"- {s}")
    return "\n".join(lines)


def collect_ai_analysis_flags(items: list[StockItem]) -> dict[str, Any]:
    """AI 프롬프트와 동일한 기준으로 고갈·가속·우선순위·환산액 플래그를 산출."""
    deplete_codes: list[str] = []
    surge_codes: list[str] = []
    by_code: dict[str, dict[str, Any]] = {}
    manufacture = select_manufacture_candidates(items)
    monitoring = select_monitoring_targets(items)
    long_term_low = select_long_term_low_items(items)
    categories = group_by_depletion_category(items)
    valuation = compute_inventory_valuation(items)

    for it in items_for_ai_analysis(items):
        stats = estimate_depletion(it)
        flag = {
            "label": it.label,
            "manage_no": it.manage_no,
            "speed": stats["speed"],
            "years_left": stats["years_left"],
            "deplete_within_2y": bool(stats["deplete_within_2y"]),
            "deplete_within_5y": bool(stats["deplete_within_5y"]),
            "recent_surge": bool(stats["recent_surge"]),
            "deplete_ym": stats["deplete_ym"],
            "depletion_category": stats["depletion_category"],
            "reliability": stats["reliability"],
            "priority_score": stats["priority_score"],
            "annual_rate": stats["annual_rate"],
            "rate_change_ratio": stats["rate_change_ratio"],
            "acceleration": stats["acceleration"],
            "acceleration_ratio": stats["acceleration_ratio"],
            "long_term_low": stats["long_term_low"],
            "stock_value": stats.get("stock_value"),
            "unit_price": stats.get("unit_price"),
            "increase_segments_excluded": stats.get("increase_segments_excluded", 0),
        }
        if it.manage_no:
            by_code[it.manage_no] = flag
        if flag["deplete_within_5y"]:
            deplete_codes.append(it.manage_no or it.label)
        if flag["acceleration"] in ("급가속", "증가"):
            surge_codes.append(it.manage_no or it.label)

    flags = {
        "by_code": by_code,
        "deplete_codes": [c for c in deplete_codes if c],
        "surge_codes": [c for c in surge_codes if c],
        "deplete_labels": [
            by_code[c]["label"] for c in deplete_codes if c in by_code
        ],
        "surge_labels": [
            by_code[c]["label"] for c in surge_codes if c in by_code
        ],
        "manufacture_candidates": manufacture,
        "monitoring_targets": monitoring,
        "long_term_low_items": long_term_low,
        "depletion_categories": categories,
        "valuation": valuation,
    }
    flags["dashboard"] = build_kpi_dashboard(items, flags)
    return flags


def extract_mentioned_codes_from_report(report: str, known_codes: list[str]) -> list[str]:
    """AI 리포트 본문에 등장한 관리번호를 추출 (긴 코드 우선)."""
    if not report:
        return []
    ordered = sorted({c for c in known_codes if c}, key=len, reverse=True)
    found: list[str] = []
    for code in ordered:
        if code in report:
            found.append(code)
    return found


def _format_candidate_lines(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "없음"
    parts = []
    for i, r in enumerate(rows, 1):
        score = r.get("priority_score")
        score_txt = f"{score:.3f}" if isinstance(score, (int, float)) else str(score)
        if i <= 3:
            risk = r.get("stock_risk")
            speed_n = r.get("speed_n")
            accel_n = r.get("accel_n")
            rel = r.get("reliability_score")
            ar = r.get("annual_rate")
            ar_txt = f"{ar:.2f}" if isinstance(ar, (int, float)) else "-"
            ratio = r.get("acceleration_ratio")
            ratio_txt = f"{ratio:.2f}" if isinstance(ratio, (int, float)) else "-"
            parts.append(
                f"{i}. {r['label']} | 점수:{score_txt} | "
                f"재고위험도({risk}) | 분양속도정규화({speed_n}) | "
                f"가속도정규화({accel_n}) | 신뢰도점수({rel}) | "
                f"연평균:{ar_txt} | 가속도:{r.get('acceleration')}/{ratio_txt} | "
                f"소진시점:{r.get('deplete_ym') or '-'}"
            )
        else:
            parts.append(
                f"{i}. {r['label']} | 점수:{score_txt} | "
                f"소진:{r.get('deplete_ym') or '-'} | "
                f"가속도:{r.get('acceleration') or '-'}"
            )
    return "\n".join(parts)


def build_ai_prompt(
    items: list[StockItem],
    *,
    compendium_context: str | None = None,
    compendium_match_report: str | None = None,
    flags: dict[str, Any] | None = None,
) -> str:
    targets = items_for_ai_analysis(items)
    bulk_dates = detect_bulk_decrease_dates(items)
    if flags is None:
        flags = collect_ai_analysis_flags(items)
    as_of = date.today()
    manufacture = flags.get("manufacture_candidates") or {}
    monitoring = flags.get("monitoring_targets") or []
    long_term_low = flags.get("long_term_low_items") or []
    categories = flags.get("depletion_categories") or {}
    valuation = flags.get("valuation") or {}
    dashboard = flags.get("dashboard") or build_kpi_dashboard(items, flags)

    lines = [
        "당신은 생약표준품 재고·분양 분석 전문가입니다.",
        "아래는 연도별 최종 재고량 기준으로 소급 보정이 완료된 데이터와 사전 정량 산출 결과입니다.",
        f"기준일(오늘): {as_of.isoformat()} ({as_of.year}년)",
        "",
        "[기본 전제]",
        "1. 생약표준품은 민간 분양에 따라 지속적으로 감소하는 것이 정상이므로, 일반적인 감소 추이는 정상으로 간주합니다.",
        "2. 재고량이 증가한 구간(전수조사/추가제조/반납 등) 및 소급 보정 증가는 정상 분양속도 계산에서 제외·보정합니다.",
        "3. 동일 날짜에 100개 이상 품목이 동시 감소한 기록은 연구 과제 목적의 대량 출고이므로 민간 분양 분석에서 제외합니다.",
        "",
        "[할루시네이션(억측) 금지 — 반드시 준수]",
        "- 제공된 재고/분양 수치·사전 산출 지표에 근거하지 않은 자의적 추측·원인 단정을 하지 마세요.",
        "- 데이터로 확인되지 않은 내용은 '확실하지 않음' 또는 '추측입니다'라고 명시하세요.",
        "- 오직 수집된 재고/분양 데이터·사전 산출 결과·(있으면) 공정서 DB만으로 객관적 사실 위주로 보고하세요.",
        "- 공정서 DB는 규격/기준 참조용이며 재고·분양 수치 산출에 사용하지 마세요.",
        "",
        "[출력 형식 — 필수]",
        "1. 리포트 맨 최상단에 아래 '사전 산출: 1페이지 요약 대시보드' 마크다운을 그대로 포함하고, "
        "자동 종합 의견을 5줄 내외로 다듬어 제시하세요.",
        "2. '분석 전문가의 제언', '전문가 제언', '제언' 섹션은 작성하지 마세요.",
        "3. 사전 산출 목록·수치를 임의로 바꾸지 마세요.",
        "",
        "[분석 요청 항목]",
        "1. 품목별 분양 속도(빠름/보통/느림) — 증가 구간 제외 속도 사용",
        "2. 소진 예상 시점을 'YYYY년 M월'로 명시하고 카테고리 목록화:",
        "   [1년 이내 (n개월)] / [1~3년 이내] / [3~5년 이내] / [5~10년 이내] / [10~15년 이내] / [15년 이상/안정]",
        "3. 데이터 신뢰도 등급 A~D (수집 횟수·관측 간격). D/부족은 '신뢰도 낮음(데이터 부족)'으로 명시",
        "4. 분양 가속도(급가속/증가/안정/감소) 및 장기 저분양·과다재고 품목(차기 제조 수량 하향 권고)",
        f"   ({ACCELERATION_FORMULA_KO})",
        "5. 차년도 제조검토대상: 제조우선순위점수 상위 · 5년 이내 소진 표준생약·지표성분 각 10종 순위표",
        "   ※ 제조 우선순위 섹션 서두에 아래 산출 공식·가중치를 그대로 인용해 명시할 것:",
        f"   {PRIORITY_FORMULA_KO}",
        "6. 모니터링 대상: 분양 가속도 급증 품목 + 정량 지표",
        "7. 현 재고 기준 분양금액 환산액(현재재고×가격): 전체·유형별·TOP20",
        "",
        f"[연구과제 대량출고로 제외할 날짜] {', '.join(bulk_dates) if bulk_dates else '해당 없음'}",
        "",
        f"[분양 가속도 산출식] {ACCELERATION_FORMULA_KO}",
        "",
        "[제조우선순위점수 산출식·가중치]",
        PRIORITY_FORMULA_KO,
        "",
    ]

    if compendium_context and compendium_context.strip():
        lines.append(compendium_context.strip())
        lines.append("")

    if compendium_match_report and compendium_match_report.strip():
        lines.append(compendium_match_report.strip())
        lines.append("")

    lines.extend([
        "[사전 산출: 1페이지 요약 대시보드]",
        format_kpi_dashboard_markdown(dashboard),
        "",
        "[사전 산출: 소진 기간 카테고리]",
    ])

    for cat_name, labels in categories.items():
        if labels:
            lines.append(f"- {cat_name} ({len(labels)}건): {', '.join(labels)}")
        else:
            lines.append(f"- {cat_name}: 없음")

    lines.append("")
    lines.append("[사전 산출: 차년도 제조검토대상 — 표준생약]")
    lines.append(_format_candidate_lines(manufacture.get("표준생약") or []))
    lines.append("")
    lines.append("[사전 산출: 차년도 제조검토대상 — 지표성분]")
    lines.append(_format_candidate_lines(manufacture.get("지표성분") or []))
    lines.append("")
    lines.append("[사전 산출: 모니터링 대상(분양 가속)]")
    surge_mon = [r for r in monitoring if r.get("acceleration") == "급가속"]
    other_mon = [r for r in monitoring if r.get("acceleration") != "급가속"]
    if surge_mon:
        lines.append(f"- 급가속 전량 ({len(surge_mon)}건):")
        for i, r in enumerate(surge_mon, 1):
            ar = r.get("annual_rate")
            ar_txt = f"{ar:.2f}" if isinstance(ar, (int, float)) else "-"
            ratio = r.get("acceleration_ratio")
            ratio_txt = f"{ratio:.2f}" if isinstance(ratio, (int, float)) else "-"
            lines.append(
                f"{i}. {r['label']} | 가속도:급가속 | 비율:{ratio_txt} | "
                f"연평균:{ar_txt} | 소진:{r.get('deplete_ym') or '-'} | "
                f"{r.get('reliability')}"
            )
    else:
        lines.append("- 급가속: 없음")
    if other_mon:
        lines.append(f"- 증가 등 ({len(other_mon)}건):")
        for i, r in enumerate(other_mon, 1):
            ar = r.get("annual_rate")
            ar_txt = f"{ar:.2f}" if isinstance(ar, (int, float)) else "-"
            lines.append(
                f"{i}. {r['label']} | {r.get('indicators')} | "
                f"연평균:{ar_txt} | 소진:{r.get('deplete_ym') or '-'} | "
                f"{r.get('reliability')}"
            )

    lines.append("")
    lines.append("[사전 산출: 장기 저분양/과다재고 — 차기 제조 수량 하향 권고]")
    if long_term_low:
        for i, r in enumerate(long_term_low[:25], 1):
            lines.append(
                f"{i}. {r['label']} | 연평균:{r['annual_rate'] if r['annual_rate'] is not None else '-'} | "
                f"환산:{_fmt_money(r.get('stock_value'))} | {r['recommendation']}"
            )
    else:
        lines.append("없음")

    lines.append("")
    lines.append("[사전 산출: 현 재고 기준 분양금액 환산액]")
    lines.append(f"- 총액: {_fmt_money(valuation.get('total_value'))}")
    for tname, tval in (valuation.get("by_type") or {}).items():
        lines.append(f"- {tname}: {_fmt_money(tval)}")
    lines.append(
        f"- 가격 반영 {valuation.get('priced_count', 0)}종 / "
        f"가격 미기재 {valuation.get('missing_price_count', 0)}종"
    )
    lines.append("- TOP20:")
    top20 = valuation.get("top20") or []
    if top20:
        for i, r in enumerate(top20, 1):
            lines.append(
                f"  {i}. {r['label']} | 재고:{r['qty']:g} × 단가:{r['unit_price']:g} = "
                f"{_fmt_money(r['stock_value'])}"
            )
    else:
        lines.append("  없음 (가격 컬럼 없음)")

    lines.append("")
    lines.append("[재고 변동 핵심 데이터]")

    for it in targets:
        stats = flags["by_code"].get(it.manage_no) or estimate_depletion(it)
        if "reliability" not in stats:
            stats = estimate_depletion(it)
        history = ", ".join(
            f"{format_year(p.change_date)}={p.quantity:g}" for p in it.corrected_points
        )
        years_left = stats.get("years_left")
        years_txt = f"{years_left:.1f}년" if years_left is not None else "산출불가/감소없음"
        rel = stats.get("reliability") or {}
        rel_label = rel.get("label", "") if isinstance(rel, dict) else str(rel)
        rel_n = rel.get("count", "?") if isinstance(rel, dict) else "?"
        lines.append(
            f"- {it.label} | 구분:{it.std_type} | 분양여부:{_cell_str(it.distributed)} | "
            f"최초:{it.first_qty:g} → 최종:{it.last_qty:g} (Δ{it.qty_delta:+g}) | "
            f"분양속도:{stats.get('speed')}(증가구간제외{stats.get('increase_segments_excluded', 0)}) | "
            f"가속도:{stats.get('acceleration')} | 예상소진:{years_txt} "
            f"({stats.get('deplete_ym') or '-'}) | 구간:{stats.get('depletion_category')} | "
            f"5년이내소진:{'예' if stats.get('deplete_within_5y') else '아니오'} | "
            f"장기저분양:{'예' if stats.get('long_term_low') else '아니오'} | "
            f"우선순위:{stats.get('priority_score')} | "
            f"환산:{_fmt_money(stats.get('stock_value'))} | "
            f"신뢰도:{rel_label}(수집{rel_n}회) | 추이:[{history}]"
        )

    lines.append("")
    lines.append(
        "위 전제·할루시네이션 금지·사전 산출·출력 형식을 반영한 한국어 마크다운 분석 리포트를 작성해 주세요. "
        "최상단 KPI 대시보드와 사실 기반 본문만 포함하고, 분석 전문가의 제언 섹션은 넣지 마세요."
    )
    return "\n".join(lines)



def find_items_by_partial_query(
    items: list[StockItem],
    question: str,
) -> list[StockItem]:
    """질문 문자열로 관리번호·한글명·영문명 부분일치(contains) 품목 검색.

    예: "탄시논" → "탄시논 IIA" 매칭.
    """
    q = (question or "").strip()
    if not q or not items:
        return []
    q_lower = q.lower()
    tokens = [t for t in re.split(r"[\s,./?!~·\-_:;|()\[\]{}]+", q) if len(t) >= 2]

    results: list[StockItem] = []
    seen: set[str] = set()
    for it in items:
        key = it.manage_no or it.label
        if key in seen:
            continue
        code = (it.manage_no or "").strip()
        ko = (it.name_ko or "").strip()
        en = ""
        if not _is_empty(getattr(it, "name_en", None)):
            en = _cell_str(it.name_en).strip()
        en_l = en.lower()
        label = (it.label or "").strip()

        hit = False
        if code and code.lower() in q_lower:
            hit = True
        elif ko and (ko in q or any(t in ko for t in tokens)):
            hit = True
        elif en and (en_l in q_lower or any(t.lower() in en_l for t in tokens)):
            hit = True
        elif label and (any(t in label for t in tokens) or label.lower() in q_lower):
            hit = True
        if not hit and tokens:
            for t in tokens:
                tl = t.lower()
                if (ko and t in ko) or (en and tl in en_l) or (code and tl in code.lower()):
                    hit = True
                    break
        if hit:
            seen.add(key)
            results.append(it)
    return results


def serialize_flags_snapshot(flags: dict[str, Any] | None) -> str:
    """챗봇용 사전 산출 스냅샷 직렬화 — 표준 리포트와 동일 수치."""
    if not flags:
        return "[사전 산출 스냅샷] 없음 (플래그 미제공)"

    by_code = flags.get("by_code") or {}
    dashboard = flags.get("dashboard") or {}
    valuation = flags.get("valuation") or {}
    categories = flags.get("depletion_categories") or {}
    manufacture = flags.get("manufacture_candidates") or {}
    monitoring = flags.get("monitoring_targets") or []
    long_term_low = flags.get("long_term_low_items") or []

    lines = [
        "[사전 산출 스냅샷 — 표준 분석 리포트와 동일 수치. 이 스냅샷·실시간 조회 수치만 사용하세요]",
        f"- by_code 품목 수: {len(by_code)}",
        f"- 5년 이내 소진 후보: {len(flags.get('deplete_codes') or [])}건",
        f"- 가속(급가속/증가) 후보: {len(flags.get('surge_codes') or [])}건",
        format_kpi_dashboard_markdown(dashboard) if dashboard else "- KPI 대시보드: 없음",
        "",
        f"- 환산 총액: {_fmt_money(valuation.get('total_value'))}",
    ]
    for tname, tval in (valuation.get("by_type") or {}).items():
        lines.append(f"  · {tname}: {_fmt_money(tval)}")

    lines.append("")
    lines.append("[스냅샷: 소진 기간 카테고리]")
    for cat_name, labels in categories.items():
        if labels:
            preview = ", ".join(labels[:12])
            more = f" 외 {len(labels) - 12}건" if len(labels) > 12 else ""
            lines.append(f"- {cat_name} ({len(labels)}건): {preview}{more}")
        else:
            lines.append(f"- {cat_name}: 없음")

    lines.append("")
    lines.append("[스냅샷: 차년도 제조검토 — 표준생약]")
    lines.append(_format_candidate_lines(manufacture.get("표준생약") or []))
    lines.append("[스냅샷: 차년도 제조검토 — 지표성분]")
    lines.append(_format_candidate_lines(manufacture.get("지표성분") or []))

    lines.append("")
    lines.append(f"[스냅샷: 모니터링 대상] {len(monitoring)}건")
    for i, r in enumerate(monitoring[:15], 1):
        lines.append(
            f"{i}. {r.get('label')} | {r.get('acceleration')} | "
            f"소진:{r.get('deplete_ym') or '-'} | 우선:{r.get('priority_score')}"
        )

    if long_term_low:
        lines.append("")
        lines.append(f"[스냅샷: 장기 저분양] {len(long_term_low)}건")
        for i, r in enumerate(long_term_low[:10], 1):
            lines.append(f"{i}. {r.get('label')} | 연평균:{r.get('annual_rate')}")

    lines.append("")
    lines.append("[스냅샷: 품목별 핵심 지표]")
    for i, (code, st) in enumerate(by_code.items()):
        if i >= 80:
            lines.append(f"... (이하 {len(by_code) - 80}종 생략, by_code에 존재)")
            break
        yl = st.get("years_left")
        yl_txt = f"{yl:.2f}" if isinstance(yl, (int, float)) else "-"
        rel = st.get("reliability") or {}
        rel_g = rel.get("grade") if isinstance(rel, dict) else None
        lines.append(
            f"- {code} | {st.get('label')} | 속도:{st.get('speed')} | "
            f"가속:{st.get('acceleration')} | 소진년:{yl_txt} | "
            f"소진월:{st.get('deplete_ym')} | f:{st.get('priority_score')} | "
            f"환산:{_fmt_money(st.get('stock_value'))} | 신뢰:{rel_g}"
        )

    text = "\n".join(lines)
    if len(text) > 28_000:
        text = text[:27_960] + "\n... (스냅샷 이하 생략)"
    return text


def _stats_from_flags(
    it: StockItem,
    by_code: dict[str, Any],
    *,
    allow_estimate: bool = True,
) -> dict[str, Any]:
    """flags by_code 스냅샷 우선, 없을 때만 estimate_depletion."""
    if it.manage_no and it.manage_no in by_code:
        return by_code[it.manage_no]
    if allow_estimate:
        return estimate_depletion(it)
    return {}


def query_live_inventory_context(
    items: list[StockItem],
    question: str,
    *,
    table_df: pd.DataFrame | None = None,
    max_detail_items: int = 35,
    flags: dict[str, Any] | None = None,
) -> str:
    """사용자 질문에 맞춰 재고 데이터를 필터링한 컨텍스트.

    flags가 주어지면 collect_ai_analysis_flags를 재호출하지 않고 스냅샷을 재사용한다.
    """
    if not items:
        return "[실시간 재고 조회] 업로드된 재고 데이터가 없습니다."

    q = (question or "").strip()
    q_lower = q.lower()
    reuse_flags = flags is not None
    if flags is None:
        flags = collect_ai_analysis_flags(items)
    dashboard = flags.get("dashboard") or build_kpi_dashboard(items, flags)
    valuation = flags.get("valuation") or compute_inventory_valuation(items)
    by_code = flags.get("by_code") or {}

    matched_items = find_items_by_partial_query(items, q)
    hit_codes = sorted(
        {it.manage_no for it in matched_items if it.manage_no},
        key=len,
        reverse=True,
    )
    hit_names = sorted(
        {it.name_ko for it in matched_items if it.name_ko},
        key=len,
        reverse=True,
    )
    codes = {it.manage_no for it in items if it.manage_no}
    names = {it.name_ko for it in items if it.name_ko}
    for c in codes:
        if c and c.lower() in q_lower and c not in hit_codes:
            hit_codes.append(c)
    for n in names:
        if n and n in q and n not in hit_names:
            hit_names.append(n)
    specific_hits = bool(matched_items or hit_codes or hit_names)

    want_deplete = any(k in q for k in ("고갈", "소진", "제조", "우선", "5년", "2년"))
    want_accel = any(k in q for k in ("가속", "급증", "모니터링", "속도"))
    want_low = any(k in q for k in ("저분양", "과다", "하향"))
    want_value = any(k in q for k in ("환산", "가격", "금액", "자산", "가치"))
    want_reliab = any(k in q for k in ("신뢰", "데이터 부족", "등급"))

    selected: list[StockItem] = []
    seen: set[str] = set()

    def _add(it: StockItem) -> None:
        key = it.manage_no or it.label
        if key in seen:
            return
        seen.add(key)
        selected.append(it)

    for it in matched_items:
        _add(it)
    if not matched_items:
        for it in items:
            if it.manage_no in hit_codes or it.name_ko in hit_names:
                _add(it)

    matched_items = list(selected)

    if not specific_hits:
        if want_deplete:
            for code in (flags.get("deplete_codes") or [])[:20]:
                it = next((x for x in items if x.manage_no == code), None)
                if it:
                    _add(it)
        if want_accel:
            for row in (flags.get("monitoring_targets") or [])[:20]:
                it = next((x for x in items if x.manage_no == row.get("manage_no")), None)
                if it:
                    _add(it)
        if want_low:
            for row in (flags.get("long_term_low_items") or [])[:20]:
                it = next((x for x in items if x.manage_no == row.get("manage_no")), None)
                if it:
                    _add(it)

    if not selected:
        def _score(it: StockItem) -> float:
            st = _stats_from_flags(it, by_code, allow_estimate=not reuse_flags)
            try:
                return float(st.get("priority_score") or 0)
            except (TypeError, ValueError):
                return 0.0

        scored = sorted(items_for_ai_analysis(items), key=_score, reverse=True)
        for it in scored[:max_detail_items]:
            _add(it)

    if specific_hits and matched_items:
        extras = [it for it in selected if it not in matched_items]
        selected = matched_items + extras[: max(0, max_detail_items - len(matched_items))]
    else:
        selected = selected[:max_detail_items]

    lines = [
        "[실시간 재고 데이터 재검토 결과 — 초기 리포트보다 이 수치를 우선하세요]",
        "아래는 코드로 계산된 정량 팩트만입니다. 이 수치 외 추론하지 마세요.",
        (
            "- 수치 출처: 사전 산출 스냅샷(by_code) 재사용"
            if reuse_flags
            else "- 수치 출처: 실시간 collect_ai_analysis_flags"
        ),
        f"- 기준 시각: {date.today().isoformat()}",
        f"- 대상 품목 {len(items)}종 · 질문 매칭/선별 상세 {len(selected)}종",
        format_kpi_dashboard_markdown(dashboard),
        "",
    ]

    if hit_codes or hit_names:
        lines.append(
            f"- 질문에서 식별된 키: 관리번호={', '.join(hit_codes) or '없음'} / "
            f"한글명={', '.join(hit_names) or '없음'}"
        )

    if want_value or specific_hits:
        lines.append(f"- 환산 총액: {_fmt_money(valuation.get('total_value'))}")
        for tname, tval in (valuation.get("by_type") or {}).items():
            lines.append(f"  · {tname}: {_fmt_money(tval)}")

    if want_reliab:
        grade_counts: dict[str, int] = defaultdict(int)
        for it in items:
            st = _stats_from_flags(it, by_code, allow_estimate=not reuse_flags)
            rel = st.get("reliability") or {}
            g = rel.get("grade", "?") if isinstance(rel, dict) else "?"
            grade_counts[str(g)] += 1
        lines.append(
            "- 신뢰도 등급 분포: "
            + ", ".join(f"{g}={n}" for g, n in sorted(grade_counts.items()))
        )

    if matched_items:
        lines.append("")
        lines.append("[질문 매칭 품목 — 정량 팩트 전체 (절단 없음)]")
        for it in matched_items:
            stats = _stats_from_flags(it, by_code, allow_estimate=True)
            if it.unit_price is None:
                it.unit_price = _extract_unit_price_from_meta(it.extra_meta)
            hist = ", ".join(
                f"{format_year(p.change_date)}={p.quantity:g}" for p in it.corrected_points
            )
            yl = stats.get("years_left")
            yl_txt = f"{yl:.4f}년" if isinstance(yl, (int, float)) else "산출불가"
            rel = stats.get("reliability") or {}
            lines.append(
                f"- {it.label} | manage_no={it.manage_no} | std_type={it.std_type} | "
                f"last_qty={it.last_qty} | unit_price={it.unit_price} | "
                f"stock_value={_fmt_money(it.stock_value)} | "
                f"speed={stats.get('speed')} | annual_rate={stats.get('annual_rate')} | "
                f"past_rate={stats.get('past_rate')} | recent_rate={stats.get('recent_rate')} | "
                f"acceleration={stats.get('acceleration')} | "
                f"acceleration_ratio={stats.get('acceleration_ratio')} | "
                f"years_left={yl_txt} | deplete_ym={stats.get('deplete_ym')} | "
                f"depletion_category={stats.get('depletion_category')} | "
                f"deplete_within_2y={stats.get('deplete_within_2y')} | "
                f"deplete_within_5y={stats.get('deplete_within_5y')} | "
                f"priority_score={stats.get('priority_score')} | "
                f"stock_risk={stats.get('stock_risk')} | "
                f"speed_n={stats.get('speed_n')} | accel_n={stats.get('accel_n')} | "
                f"reliability={rel.get('label') if isinstance(rel, dict) else rel} | "
                f"reliability_grade={rel.get('grade') if isinstance(rel, dict) else None} | "
                f"increase_segments_excluded={stats.get('increase_segments_excluded', 0)} | "
                f"추이:[{hist}]"
            )

    lines.append("")
    lines.append("[선별 품목 실시간 지표]")
    for it in selected:
        if matched_items and it in matched_items:
            continue
        stats = _stats_from_flags(it, by_code, allow_estimate=not reuse_flags)
        if not stats:
            continue
        hist = ", ".join(
            f"{format_year(p.change_date)}={p.quantity:g}" for p in it.corrected_points
        )
        yl = stats.get("years_left")
        yl_txt = f"{yl:.2f}년" if isinstance(yl, (int, float)) else "산출불가"
        rel = stats.get("reliability") or {}
        lines.append(
            f"- {it.label} | {it.std_type} | 최종:{it.last_qty} | "
            f"속도:{stats.get('speed')} | 연평균:{stats.get('annual_rate')} | "
            f"가속도:{stats.get('acceleration')} | 소진:{yl_txt}({stats.get('deplete_ym')}) | "
            f"구간:{stats.get('depletion_category')} | 우선순위:{stats.get('priority_score')} | "
            f"환산:{_fmt_money(stats.get('stock_value'))} | "
            f"신뢰도:{(rel.get('label') if isinstance(rel, dict) else rel)} | "
            f"증가구간제외:{stats.get('increase_segments_excluded', 0)} | 추이:[{hist}]"
        )

    if table_df is not None and not table_df.empty and (hit_codes or hit_names or matched_items):
        lines.append("")
        lines.append("[통합 표(DataFrame) 매칭 행]")
        df = table_df
        mask = None
        if "관리번호" in df.columns and hit_codes:
            mask = df["관리번호"].astype(str).isin(hit_codes)
        match_names = hit_names or [it.name_ko for it in matched_items if it.name_ko]
        if "한글명" in df.columns and match_names:
            name_mask = df["한글명"].astype(str).isin(match_names)
            for n in match_names:
                name_mask = name_mask | df["한글명"].astype(str).str.contains(
                    re.escape(n), regex=True, na=False
                )
            mask = name_mask if mask is None else (mask | name_mask)
        if mask is not None:
            sub = df.loc[mask]
            prefer = [
                c for c in (
                    "관리번호", "한글명", "영문명", "표준품구분", "재고", "잔고",
                    "가격", "가격(원)", "분양여부",
                )
                if c in df.columns
            ]
            other = [c for c in df.columns if c not in prefer and not str(c).startswith("_")]
            cols = prefer + other
            for _, row in sub.iterrows():
                brief = []
                for col in cols:
                    v = row.get(col)
                    if _is_empty(v):
                        continue
                    brief.append(f"{col}={_cell_str(v)}")
                lines.append("- " + " | ".join(brief))

    text = "\n".join(lines)
    if len(text) > 24_000 and not matched_items:
        text = text[:23_960] + "\n... (이하 생략)"
    elif len(text) > 40_000:
        text = text[:39_960] + "\n... (이하 생략)"
    return text


def build_followup_prompt(
    base_report: str,
    user_question: str,
    items: list[StockItem] | None = None,
    *,
    compendium_context: str | None = None,
    table_df: pd.DataFrame | None = None,
    flags: dict[str, Any] | None = None,
) -> str:
    """초기 리포트 이후 추가 질문용 — 사전 산출 스냅샷과 실시간 조회를 주입."""
    lines = [
        "당신은 생약표준품 재고·분양 분석 전문가입니다.",
        "사용자의 후속 질문에는 초기 리포트 문구만 반복하지 말고, "
        "아래 [사전 산출 스냅샷]과 [실시간 재고 데이터 재검토 결과]·공정서 DB를 우선 근거로 답하세요.",
        "",
        "[할루시네이션 금지] 제공된 스냅샷·실시간 수치·공정서 DB에 없는 내용을 단정하지 마세요. "
        "불확실하면 '확실하지 않음' 또는 '추측입니다'라고 밝히세요.",
        "[형식] '분석 전문가의 제언' 섹션은 작성하지 마세요.",
        "[공정서 DB] 규격/기준 참조에만 사용하고, 재고·분양 수치 산출에는 쓰지 마세요.",
        "[수치 동기화] 표준 분석 리포트 생성 시 산출된 스냅샷 수치만 사용하세요. "
        "임의로 다시 계산·추정하지 마세요.",
        "[우선순위] 초기 리포트 문구와 스냅샷/실시간 수치가 다르면 스냅샷·실시간 수치를 사용하세요.",
        "아래는 코드로 계산된 정량 팩트만입니다. 이 수치 외 추론하지 마세요.",
        "",
    ]

    if flags:
        lines.append(serialize_flags_snapshot(flags))
        lines.append("")
    else:
        lines.append("[사전 산출 스냅샷] 미제공 — 실시간 조회 결과만 사용")
        lines.append("")

    if items:
        lines.append(
            query_live_inventory_context(
                items, user_question, table_df=table_df, flags=flags
            )
        )
        lines.append("")
    else:
        lines.append("[실시간 재고 데이터 재검토 결과] 재고 데이터 없음")
        lines.append("")

    if compendium_context and compendium_context.strip():
        ctx = compendium_context.strip()
        if len(ctx) > 8000:
            ctx = ctx[:7980] + "\n... (이하 생략)"
        lines.append(ctx)
        lines.append("")

    report = (base_report or "").strip()
    if report:
        if len(report) > 3500:
            report = report[:3480] + "\n... (초기 리포트 일부만 첨부)"
        lines.append("[초기 분석 리포트 — 보조 참고]")
        lines.append(report)
        lines.append("")

    lines.append(f"[사용자 질문]\n{user_question.strip()}")
    lines.append("")
    lines.append(
        "질문에 대해 한국어 마크다운으로 답하되, 필요한 경우 스냅샷·실시간 수치(연도·수량·소진시점·"
        "가속도·환산액·신뢰도 등급·우선순위점수 등)를 명시해 주세요. "
        "스냅샷에 있는 수치 외 추론·재계산은 하지 마세요."
    )
    return "\n".join(lines)


def _md_inline_to_html(text: str) -> str:
    """Escape + simple **bold** / `code` inline markdown."""
    esc = html.escape(text)
    esc = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", esc)
    esc = re.sub(r"`([^`]+)`", r"<code>\1</code>", esc)
    return esc


def _collapse_comma_items_html(body: str, preview_n: int) -> str:
    """긴 쉼표 구분 품목 목록을 details로 접기."""
    parts = [p.strip() for p in body.split(",") if p.strip()]
    if len(parts) <= preview_n:
        return _md_inline_to_html(body)
    preview = ", ".join(parts[:preview_n])
    rest = ", ".join(parts[preview_n:])
    n = len(parts)
    return (
        f"{_md_inline_to_html(preview)}, … "
        f"<details style=\"display:inline;margin-left:4px;\">"
        f"<summary style=\"cursor:pointer;color:#1e3a5f;\">전체 {n}개 품목 펼쳐보기</summary>"
        f"<div style=\"margin-top:6px;line-height:1.55;\">{_md_inline_to_html(rest)}</div>"
        f"</details>"
    )


def markdown_report_to_collapsible_html(md_text: str, preview_n: int = 8) -> str:
    """마크다운 리포트를 HTML로 변환하고, 긴 쉼표 품목 목록은 접기 UI로 감싼다."""
    if not md_text:
        return ""
    out: list[str] = []
    in_ul = False

    def _close_ul() -> None:
        nonlocal in_ul
        if in_ul:
            out.append("</ul>")
            in_ul = False

    for raw in md_text.splitlines():
        line = raw.rstrip()
        if not line.strip():
            _close_ul()
            out.append("<br/>")
            continue

        hm = re.match(r"^(#{1,4})\s+(.*)$", line.strip())
        if hm:
            _close_ul()
            level = len(hm.group(1))
            out.append(f"<h{level}>{_md_inline_to_html(hm.group(2))}</h{level}>")
            continue

        lm = re.match(r"^(\s*[-*+]|\s*\d+\.)\s+(.*)$", line)
        if lm:
            if not in_ul:
                out.append("<ul>")
                in_ul = True
            content = lm.group(2)
            if ":" in content and content.count(",") >= preview_n:
                prefix, _, rest = content.partition(":")
                if rest.strip():
                    inner = (
                        f"{_md_inline_to_html(prefix)}: "
                        f"{_collapse_comma_items_html(rest.strip(), preview_n)}"
                    )
                else:
                    inner = _md_inline_to_html(content)
            elif content.count(",") >= max(preview_n, 8):
                inner = _collapse_comma_items_html(content, preview_n)
            else:
                inner = _md_inline_to_html(content)
            out.append(f"<li>{inner}</li>")
            continue

        _close_ul()
        if line.count(",") >= max(preview_n, 8) and ":" in line:
            prefix, _, rest = line.partition(":")
            if rest.strip() and rest.count(",") >= preview_n - 1:
                out.append(
                    f"<p>{_md_inline_to_html(prefix)}: "
                    f"{_collapse_comma_items_html(rest.strip(), preview_n)}</p>"
                )
                continue
        out.append(f"<p>{_md_inline_to_html(line)}</p>")

    _close_ul()
    style = (
        "<div style=\"font-family:'Malgun Gothic','Segoe UI',sans-serif;"
        "font-size:10.5pt;line-height:1.65;color:#1a2332;padding:4px;\">"
    )
    return style + "\n".join(out) + "</div>"



def category_counts(items: list[StockItem]) -> dict[str, int]:
    counts: dict[str, int] = defaultdict(int)
    for it in items:
        counts[it.std_type or "미분류"] += 1
    return dict(counts)


# ---------------------------------------------------------------------------
# 공정서(규격) 참조 DB — 재고 시계열 파싱 대상 아님, AI 프롬프트 컨텍스트 전용
# ---------------------------------------------------------------------------

COMPENDIUM_NAME_KEYS = (
    "생약명(한글)", "생약명", "한글명", "품목명", "명칭", "표준품명", "name", "Name",
)
COMPENDIUM_NAME_EN_KEYS = (
    "생약명(영어)", "생약명(영문)", "영문명", "영문", "영어명", "name_en", "Name_EN", "English",
)
COMPENDIUM_ORIGIN_KO_KEYS = ("기원(한글)", "기원", "기원식물", "원식물", "origin")
COMPENDIUM_ORIGIN_EN_KEYS = ("기원(영어)", "기원(영문)", "origin_en", "Origin_EN")
COMPENDIUM_PHARMACOPOEIA_KEYS = ("공정서", "수재공정서", "pharmacopoeia", "Pharmacopoeia", "KP", "KHP")
COMPENDIUM_CODE_KEYS = ("관리번호", "품목코드", "코드", "code", "Code")


@dataclass
class CompendiumEntry:
    name_ko: str = ""
    name_en: str = ""
    origin_ko: str = ""
    origin_en: str = ""
    pharmacopoeia: str = ""
    raw: dict[str, Any] = field(default_factory=dict)


def _norm_key(s: Any) -> str:
    """소문자·공백·구두점 제거 정규화 키."""
    if _is_empty(s):
        return ""
    text = str(s).strip().lower()
    text = re.sub(r"[\s\W_]+", "", text, flags=re.UNICODE)
    return text


def _pick_column(columns: Sequence[str], keywords: Sequence[str]) -> Optional[str]:
    for key in keywords:
        for col in columns:
            if key.lower() in str(col).lower():
                return str(col)
    return None


def _cell_field(row: pd.Series, col: Optional[str]) -> str:
    if not col:
        return ""
    val = row.get(col)
    if _is_empty(val):
        return ""
    return _cell_str(val).strip()


def _parse_compendium_entries(df: pd.DataFrame) -> list[CompendiumEntry]:
    columns = [str(c) for c in df.columns]
    col_ko = _pick_column(columns, COMPENDIUM_NAME_KEYS)
    col_en = _pick_column(columns, COMPENDIUM_NAME_EN_KEYS)
    col_orig_ko = _pick_column(columns, COMPENDIUM_ORIGIN_KO_KEYS)
    col_orig_en = _pick_column(columns, COMPENDIUM_ORIGIN_EN_KEYS)
    col_pharm = _pick_column(columns, COMPENDIUM_PHARMACOPOEIA_KEYS)

    entries: list[CompendiumEntry] = []
    for _, row in df.iterrows():
        name_ko = _cell_field(row, col_ko)
        name_en = _cell_field(row, col_en)
        if not name_ko and not name_en:
            continue
        raw = {str(c): row.get(c) for c in columns if not str(c).startswith("_")}
        entries.append(
            CompendiumEntry(
                name_ko=name_ko,
                name_en=name_en,
                origin_ko=_cell_field(row, col_orig_ko),
                origin_en=_cell_field(row, col_orig_en),
                pharmacopoeia=_cell_field(row, col_pharm),
                raw=raw,
            )
        )
    return entries


def load_compendium_excel(path: PathLike) -> dict[str, Any]:
    """공정서/규격 DB 엑셀을 DataFrame + 구조화 entries로 로드 (시계열·소급보정 없음)."""
    path = str(path)
    frames: list[pd.DataFrame] = []
    with pd.ExcelFile(path, engine="openpyxl") as xls:
        for sheet in xls.sheet_names:
            part = pd.read_excel(xls, sheet_name=sheet)
            part.columns = [str(c).strip() for c in part.columns]
            if part.empty:
                continue
            part = part.copy()
            part.insert(0, "_시트", sheet)
            frames.append(part)
    if not frames:
        raise ValueError(f"공정서 DB 시트에 데이터가 없습니다: {Path(path).name}")

    df = pd.concat(frames, ignore_index=True, sort=False)
    entries = _parse_compendium_entries(df)
    return {
        "file_path": path,
        "file_name": Path(path).name,
        "dataframe": df,
        "row_count": int(len(df)),
        "columns": [str(c) for c in df.columns],
        "sheet_count": len(frames),
        "entries": entries,
        "parsed_count": len(entries),
    }


def _pharmacopoeia_tag_from_text(pharm: str) -> str:
    """공정서 문자열 → '[KP 수재]' / '[KHP(생약규격집) 수재]'."""
    if not pharm:
        return ""
    text = str(pharm).strip()
    upper = text.upper()
    # 약전외·생약규격집 = KHP (KP와 혼동 방지: '약전외'를 먼저 판정)
    is_khp = (
        "KHP" in upper
        or "생약규격집" in text
        or "약전외" in text
        or "한약(생약)규격" in text
        or "한약규격집" in text
    )
    is_kp = (
        bool(re.search(r"(?<![A-Z])KP(?![A-Z])", upper))
        or ("대한민국약전" in text and "약전외" not in text)
        or text in ("약전", "KP", "kp")
    )
    tags: list[str] = []
    if is_kp:
        tags.append("KP")
    if is_khp:
        tags.append("KHP(생약규격집)")
    if not tags:
        m = re.search(r"\b([A-Z]{2,5})\b", upper)
        if m and m.group(1) not in ("IIA", "IIB", "III"):
            tags.append(m.group(1))
    if not tags:
        return f"[{text} 수재]" if text else ""
    return " ".join(f"[{t} 수재]" for t in tags)


def match_compendium_inventory(
    entries: list[CompendiumEntry],
    items: list[StockItem],
) -> dict[str, Any]:
    """공정서 entries ↔ 재고 품목 매칭 (재고 엑셀 데이터는 변경하지 않음)."""
    corrections: list[dict[str, Any]] = []
    by_manage_no: dict[str, str] = {}
    by_label: dict[str, str] = {}
    matched_entry_ids: set[int] = set()

    # 인덱스: exact ko / norm ko / norm en
    by_exact_ko: dict[str, list[CompendiumEntry]] = defaultdict(list)
    by_norm_ko: dict[str, list[CompendiumEntry]] = defaultdict(list)
    by_norm_en: dict[str, list[CompendiumEntry]] = defaultdict(list)
    by_norm_origin_ko: dict[str, list[CompendiumEntry]] = defaultdict(list)

    for e in entries:
        if e.name_ko:
            by_exact_ko[e.name_ko.strip()].append(e)
            nk = _norm_key(e.name_ko)
            if nk:
                by_norm_ko[nk].append(e)
        if e.name_en:
            ne = _norm_key(e.name_en)
            if ne:
                by_norm_en[ne].append(e)
        if e.origin_ko:
            no = _norm_key(e.origin_ko)
            if no:
                by_norm_origin_ko[no].append(e)

    def _pick_first(cands: list[CompendiumEntry]) -> Optional[CompendiumEntry]:
        for c in cands:
            if id(c) not in matched_entry_ids:
                return c
        return cands[0] if cands else None

    for it in items:
        stock_ko = (it.name_ko or "").strip()
        stock_en = _cell_str(it.name_en).strip() if not _is_empty(it.name_en) else ""
        hit: Optional[CompendiumEntry] = None
        match_type = ""

        if stock_ko and stock_ko in by_exact_ko:
            hit = _pick_first(by_exact_ko[stock_ko])
            match_type = "exact_ko"
        if hit is None and stock_ko:
            nk = _norm_key(stock_ko)
            if nk and nk in by_norm_ko:
                hit = _pick_first(by_norm_ko[nk])
                match_type = "fuzzy_ko"
        if hit is None and stock_en:
            ne = _norm_key(stock_en)
            if ne and ne in by_norm_en:
                hit = _pick_first(by_norm_en[ne])
                match_type = "fuzzy_en"
        if hit is None and stock_ko:
            # 기원 필드를 보조 신호로 사용 (한글명 일부가 기원에 포함되는 경우 등)
            nk = _norm_key(stock_ko)
            if nk and nk in by_norm_origin_ko:
                hit = _pick_first(by_norm_origin_ko[nk])
                match_type = "origin_ko"
            else:
                for ok, cands in by_norm_origin_ko.items():
                    if nk and (nk in ok or ok in nk) and len(nk) >= 2:
                        hit = _pick_first(cands)
                        match_type = "origin_ko"
                        break

        if hit is None:
            continue

        matched_entry_ids.add(id(hit))
        tag = _pharmacopoeia_tag_from_text(hit.pharmacopoeia)
        short = ""
        if tag:
            # "[KHP(생약규격집) 수재]" → "KHP(생약규격집)" (lookup이 다시 태그화 가능하도록)
            inners = re.findall(r"\[(.+?)\s*수재\]", tag)
            short = inners[0].strip() if inners else ""
            if not short:
                found = re.findall(r"\[([^\]]+)", tag)
                short = found[0].strip() if found else (hit.pharmacopoeia or "")
            if short.upper() == "KHP" or (
                short.upper().startswith("KHP") and "생약규격집" not in short
            ):
                short = "KHP(생약규격집)"
        elif hit.pharmacopoeia:
            short = hit.pharmacopoeia.strip()
            if "생약규격집" in short or "약전외" in short or "KHP" in short.upper():
                short = "KHP(생약규격집)"
        if it.manage_no and short:
            by_manage_no[it.manage_no] = short
        if tag:
            by_label[it.label] = tag

        corrections.append(
            {
                "stock_label": it.label,
                "stock_name_ko": stock_ko,
                "stock_name_en": stock_en,
                "matched_name_ko": hit.name_ko,
                "matched_name_en": hit.name_en,
                "matched_origin_ko": hit.origin_ko,
                "matched_origin_en": hit.origin_en,
                "pharmacopoeia": hit.pharmacopoeia,
                "match_type": match_type,
            }
        )

    missing: list[CompendiumEntry] = [
        e for e in entries if id(e) not in matched_entry_ids
    ]

    return {
        "corrections": corrections,
        "missing": missing,
        "by_manage_no": by_manage_no,
        "by_label": by_label,
        "correction_count": len(corrections),
    }


def format_compendium_match_report(match_result: dict[str, Any]) -> str:
    """보정 매칭·미보유 공정서 표준품 보고 텍스트."""
    corrections = match_result.get("corrections") or []
    missing = match_result.get("missing") or []
    lines = [
        f"[공정서 매칭 보정] 총 {len(corrections)}건 보정 매칭 완료",
    ]
    if corrections:
        for i, c in enumerate(corrections, 1):
            lines.append(
                f"{i}. {c.get('stock_label')} ← {c.get('matched_name_ko') or c.get('matched_name_en')} "
                f"(type={c.get('match_type')}) | 기원:{c.get('matched_origin_ko') or '-'} | "
                f"공정서:{c.get('pharmacopoeia') or '-'}"
            )
    else:
        lines.append("- 보정 매칭 없음")

    lines.append("")
    lines.append(f"[미보유 공정서 표준품] ({len(missing)}건)")
    if not missing:
        lines.append("- 없음")
    elif len(missing) > 40:
        names = [e.name_ko or e.name_en for e in missing[:40]]
        lines.append(f"- 총 {len(missing)}건 중 상위 40건: {', '.join(names)}")
        lines.append(f"- … 외 {len(missing) - 40}건 생략")
    else:
        for e in missing:
            lines.append(f"- {e.name_ko or e.name_en}" + (f" / {e.name_en}" if e.name_ko and e.name_en else ""))

    return "\n".join(lines)


def lookup_pharmacopoeia_tag(
    item: StockItem,
    match_result: dict[str, Any] | None,
) -> str:
    """품목에 대한 '[KP 수재]' 형태 태그 또는 빈 문자열."""
    if not match_result:
        return ""
    by_label = match_result.get("by_label") or {}
    if item.label in by_label:
        return by_label[item.label]
    by_manage = match_result.get("by_manage_no") or {}
    code = by_manage.get(item.manage_no or "")
    if code:
        return _pharmacopoeia_tag_from_text(str(code))
    return ""


def _row_to_brief(row: pd.Series, columns: list[str], max_cols: int = 12) -> str:
    parts: list[str] = []
    for col in columns[:max_cols]:
        if str(col).startswith("_"):
            continue
        val = row.get(col)
        if _is_empty(val):
            continue
        text = _cell_str(val)
        if len(text) > 80:
            text = text[:77] + "..."
        parts.append(f"{col}={text}")
    return " | ".join(parts)


def format_compendium_context(
    df: pd.DataFrame | None,
    items: list[StockItem] | None = None,
    *,
    max_rows: int = 60,
    max_chars: int = 12_000,
    meta: dict[str, Any] | None = None,
    match_result: dict[str, Any] | None = None,
) -> str:
    """AI 프롬프트용 공정서 DB 요약·매칭 행 텍스트.

    main.py 에서 df(+meta)만 넘겨도 동작. items가 있으면 구조화 매칭을 우선 사용.
    """
    if df is None or df.empty:
        return ""

    columns = [str(c) for c in df.columns]
    entries: list[CompendiumEntry] = []
    if meta and meta.get("entries"):
        entries = list(meta["entries"])
    else:
        entries = _parse_compendium_entries(df)

    lines = [
        "[공정서/규격 참조 DB — 재고 시계열이 아님. 규격·기준 정보 참조 전용]",
        f"- 파일: {(meta or {}).get('file_name', '')}",
        f"- 행 수: {len(df)} · 구조화 entries: {len(entries)} · 열: {', '.join(columns[:20])}"
        + (" ..." if len(columns) > 20 else ""),
        "- 사용 규칙: 재고/분양 수치 산출에 쓰지 말고, 품목 규격·기준 설명에만 인용하세요.",
        "- DB에 없는 규격 정보는 '공정서 DB에 해당 항목 없음'으로 명시하세요.",
    ]

    if entries:
        lines.append(
            f"- entries 요약(상위 {min(8, len(entries))}): "
            + ", ".join(
                (e.name_ko or e.name_en) + (f"/{e.pharmacopoeia}" if e.pharmacopoeia else "")
                for e in entries[:8]
            )
        )

    local_match = match_result
    if local_match is None and items and entries:
        local_match = match_compendium_inventory(entries, items)

    if local_match and (local_match.get("corrections") or local_match.get("missing") is not None):
        lines.append("")
        report = format_compendium_match_report(local_match)
        # 컨텍스트 길이 보호를 위해 리포트도 포함하되 아래에서 max_chars로 절단
        lines.append(report)

    name_col = _pick_column(columns, COMPENDIUM_NAME_KEYS)
    code_col = _pick_column(columns, COMPENDIUM_CODE_KEYS)
    keys_name = {it.name_ko.strip() for it in (items or []) if it.name_ko}
    keys_code = {it.manage_no.strip() for it in (items or []) if it.manage_no}

    def _row_matches(row: pd.Series) -> bool:
        if name_col and not _is_empty(row.get(name_col)):
            nm = _cell_str(row.get(name_col))
            if nm in keys_name or any(nm and (nm in k or k in nm) for k in keys_name):
                return True
        if code_col and not _is_empty(row.get(code_col)):
            cd = _cell_str(row.get(code_col))
            if cd in keys_code:
                return True
        return False

    matched_rows: list[pd.Series] = []
    if keys_name or keys_code:
        for _, row in df.iterrows():
            if _row_matches(row):
                matched_rows.append(row)

    lines.append("")
    if matched_rows:
        lines.append(
            f"[재고 품목과 매칭된 공정서 행] ({min(len(matched_rows), max_rows)}/{len(matched_rows)}건 표시)"
        )
        for row in matched_rows[:max_rows]:
            lines.append(f"- {_row_to_brief(row, columns)}")
    elif not local_match:
        lines.append("[공정서 DB 미리보기] (재고 품목과 자동 매칭된 행 없음 — 상위 행 샘플)")
        preview_n = min(15, len(df), max_rows)
        for _, row in df.head(preview_n).iterrows():
            lines.append(f"- {_row_to_brief(row, columns)}")

    text = "\n".join(lines)
    if len(text) > max_chars:
        text = text[: max_chars - 20] + "\n... (이하 생략)"
    return text


# 3D 표시용 예상 소진기간 상한 (AI 추이와 동일 산출, 초장기는 상한에 묶음)
YEARS_LEFT_DISPLAY_CAP = 50.0


def scatter_cat_label(std_type: str, std_type_raw: Any = None) -> str:
    """3D 범례용 표준품구분 (접두어 없이 표준생약/지표성분/대조품)."""
    text = std_type or _cell_str(std_type_raw)
    if not text:
        return "미분류"
    for key, label in (
        ("표준생약", "표준생약"),
        ("지표성분", "지표성분"),
        ("대조품", "대조품"),
        ("대조생약", "대조품"),
    ):
        if key in text:
            return label
    match = re.search(r"\(([^)]+)\)", text)
    if match:
        return match.group(1).strip() or text
    return text


def build_scatter3d_record(
    item: StockItem,
    ai_flags: dict[str, Any] | None = None,
    mentioned_codes: set[str] | None = None,
) -> Optional[dict[str, Any]]:
    """AI 추이(estimate_depletion) 기준 3D 레코드. 감소 추이가 없으면 None."""
    if not item.has_stock_change:
        return None

    pts = item.corrected_points
    if len(pts) < 2 or item.first_qty is None or item.last_qty is None:
        return None

    flag = (ai_flags or {}).get(item.manage_no) if ai_flags else None
    stats = estimate_depletion(item)
    if flag is not None:
        deplete = bool(flag.get("deplete_within_5y"))
        surge = bool(flag.get("recent_surge"))
        speed = flag.get("speed") or stats.get("speed")
        years_left = flag.get("years_left")
        if years_left is None:
            years_left = stats.get("years_left")
        annual_rate = stats.get("annual_rate")
    else:
        deplete = bool(stats.get("deplete_within_5y"))
        surge = bool(stats.get("recent_surge"))
        speed = stats.get("speed")
        years_left = stats.get("years_left")
        annual_rate = stats.get("annual_rate")

    # 감소 없음/증가만 있는 품목은 천년대 runway로 왜곡되므로 제외
    if annual_rate is None or annual_rate <= 1e-9 or years_left is None:
        return None

    years_raw = float(years_left)
    years_plot = min(years_raw, YEARS_LEFT_DISPLAY_CAP)
    capped = years_raw > YEARS_LEFT_DISPLAY_CAP

    qtys = [p.quantity for p in pts]
    total_variation = sum(abs(qtys[i + 1] - qtys[i]) for i in range(len(qtys) - 1))
    net_drop = float(item.first_qty - item.last_qty)
    init_qty = float(item.first_qty)
    balance = float(item.last_qty)
    days = (pts[-1].change_date - pts[0].change_date).days or 1
    elapsed_years = max(days / 365.25, 1 / 365.25)
    if abs(init_qty) > 1e-12:
        decrease_rate = net_drop / init_qty * 100.0
    else:
        decrease_rate = 0.0

    mentioned = bool(mentioned_codes and item.manage_no in mentioned_codes)

    return {
        "cat": scatter_cat_label(item.std_type, item.std_type_raw),
        "code": item.manage_no,
        "name": item.name_ko,
        "balance": round(balance, 6),
        "initQty": round(init_qty, 6),
        "netDrop": round(net_drop, 6),
        "totalVariation": round(float(total_variation), 6),
        "decreaseRate": round(float(decrease_rate), 4),
        "annualRate": round(float(annual_rate), 6),
        "runway": round(years_plot, 6),
        "elapsedYears": round(float(elapsed_years), 6),
        "depleteWithin5y": deplete,
        "recentSurge": surge,
        "speed": speed or "",
        "yearsLeft": round(years_plot, 4),
        "yearsLeftRaw": round(years_raw, 4),
        "yearsCapped": capped,
        "aiMentioned": mentioned,
    }


def build_scatter3d_records(
    items: list[StockItem],
    ai_flags: dict[str, Any] | None = None,
    mentioned_codes: set[str] | list[str] | None = None,
) -> list[dict[str, Any]]:
    """유효한 3D 산점도 레코드만 모은다."""
    flag_map = None
    if isinstance(ai_flags, dict) and "by_code" in ai_flags:
        flag_map = ai_flags["by_code"]
    elif isinstance(ai_flags, dict):
        flag_map = ai_flags
    mentioned = set(mentioned_codes or [])
    records: list[dict[str, Any]] = []
    for item in items:
        rec = build_scatter3d_record(item, ai_flags=flag_map, mentioned_codes=mentioned)
        if rec is not None:
            records.append(rec)
    return records
