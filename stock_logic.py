"""생약표준품 재고 엑셀 파싱 · 연도별 소급 보정 · AI 프롬프트 구성."""

from __future__ import annotations

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
MANUFACTURE_CANDIDATE_LIMIT = 20
LONG_TERM_LOW_YEARS = 5.0
LONG_TERM_LOW_REL_DROP = 0.05  # 5년+ 구간 상대 감소율 5% 미만 → 저분양
PRICE_COL_KEYWORDS = ("가격",)
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


def _half_period_decrease_rate(points: list[StockPoint]) -> tuple[Optional[float], Optional[float]]:
    """전반/후반 감소 구간 연평균 속도."""
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
    """분양 가속도: 급가속 / 증가 / 안정 / 감소."""
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
    return {"label": label, "ratio": ratio, "is_surge": label in ("급가속", "증가") and ratio >= 1.25}


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
        return f"1년 이내({months}개월 이내)"
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


def manufacturing_priority_score(
    *,
    years_left: Optional[float],
    annual_rate: Optional[float],
    acceleration_ratio: Optional[float],
    reliability_score: float,
) -> float:
    """제조우선순위점수 = f(재고위험도, 최근 분양속도, 분양 가속도, 데이터 신뢰도)."""
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
    return round(0.40 * risk + 0.25 * speed_n + 0.20 * accel_n + 0.15 * rel, 4)


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
    early_rate, late_rate = _half_period_decrease_rate(pts)
    accel = acceleration_from_rates(early_rate, late_rate)
    rate_change_ratio = accel["ratio"]
    recent_surge = bool(accel["is_surge"] and accel["label"] == "급가속")

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
    priority = manufacturing_priority_score(
        years_left=years_left,
        annual_rate=annual_rate,
        acceleration_ratio=accel["ratio"],
        reliability_score=reliability["score"],
    )
    stock_risk = 0.0
    if years_left is not None:
        if years_left <= 0:
            stock_risk = 1.0
        else:
            stock_risk = max(0.0, min(1.0, 1.0 - (years_left / 15.0)))

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
        "long_term_low": long_term_low,
        "increase_segments_excluded": dec["increase_segments"],
        "deplete_ym": deplete_ym,
        "depletion_category": category,
        "reliability": reliability,
        "stock_risk": stock_risk,
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
        bucket.append(
            {
                "label": it.label,
                "manage_no": it.manage_no,
                "std_type": it.std_type,
                "priority_score": score,
                "years_left": stats["years_left"],
                "deplete_ym": stats["deplete_ym"],
                "depletion_category": stats["depletion_category"],
                "reliability": rel["label"] if isinstance(rel, dict) else str(rel),
                "reliability_grade": rel.get("grade") if isinstance(rel, dict) else None,
                "annual_rate": stats["annual_rate"],
                "acceleration": stats["acceleration"],
            }
        )
    return result


def select_monitoring_targets(items: list[StockItem], limit: int = 30) -> list[dict[str, Any]]:
    """분양 가속도 급증(급가속/증가) 모니터링 대상."""
    rows: list[dict[str, Any]] = []
    for it in items:
        stats = estimate_depletion(it)
        if stats.get("acceleration") not in ("급가속", "증가"):
            continue
        rows.append(
            {
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
                        f", 후반/전반비={stats['acceleration_ratio']:.2f}"
                        if stats["acceleration_ratio"] is not None
                        else ""
                    )
                ),
            }
        )
    rows.sort(
        key=lambda r: (
            1 if r["acceleration"] == "급가속" else 0,
            float(r["acceleration_ratio"] or 0),
            float(r["priority_score"] or 0),
        ),
        reverse=True,
    )
    return rows[:limit]


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
    """핵심 KPI 8~10개 + 자동 종합 의견(5줄 내외)."""
    flags = flags or {}
    valuation = flags.get("valuation") or compute_inventory_valuation(items)
    by_code = flags.get("by_code") or {}

    managed = len(items)
    changed = len(items_for_ai_analysis(items))
    deplete_2y = 0
    # deplete_codes가 없으면 직접 집계
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
        {"key": "managed", "label": "관리품목 수", "value": managed, "display": f"{managed}종"},
        {"key": "changed", "label": "변동 품목 수", "value": changed, "display": f"{changed}종"},
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
        f"관리 품목 {managed}종 중 재고 변동이 확인된 품목은 {changed}종입니다.",
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
        yl = r.get("years_left")
        yl_txt = f"{yl:.1f}년" if isinstance(yl, (int, float)) else "-"
        accel = r.get("acceleration")
        accel_txt = f" | 가속도:{accel}" if accel else ""
        parts.append(
            f"{i}. {r['label']} | 점수:{r['priority_score']:.3f} | "
            f"소진:{r.get('deplete_ym') or yl_txt} | "
            f"구간:{r.get('depletion_category')} | {r.get('reliability')}{accel_txt}"
        )
    return "\n".join(parts)


def build_ai_prompt(items: list[StockItem]) -> str:
    targets = items_for_ai_analysis(items)
    bulk_dates = detect_bulk_decrease_dates(items)
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
        "- 오직 수집된 재고/분양 데이터와 아래 사전 산출 결과만으로 객관적 사실 위주로 보고하세요.",
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
        "   [1년 이내(n개월 이내)] / [1~3년 이내] / [3~5년 이내] / [5~10년 이내] / [10~15년 이내] / [15년 이상/안정]",
        "3. 데이터 신뢰도 등급 A~D (수집 횟수·관측 간격). D/부족은 '신뢰도 낮음(데이터 부족)'으로 명시",
        "4. 분양 가속도(급가속/증가/안정/감소) 및 장기 저분양·과다재고 품목(차기 제조 수량 하향 권고)",
        "5. 차년도 제조검토대상: 제조우선순위점수 상위 · 5년 이내 소진 표준생약·지표성분 각 20종 내외 순위표",
        "   (점수 = f(재고위험도, 최근 분양속도, 분양 가속도, 데이터 신뢰도))",
        "6. 모니터링 대상: 분양 가속도 급증 품목 + 정량 지표",
        "7. 현 재고 기준 분양금액 환산액(현재재고×가격): 전체·유형별·TOP20",
        "",
        f"[연구과제 대량출고로 제외할 날짜] {', '.join(bulk_dates) if bulk_dates else '해당 없음'}",
        "",
        "[사전 산출: 1페이지 요약 대시보드]",
        format_kpi_dashboard_markdown(dashboard),
        "",
        "[사전 산출: 소진 기간 카테고리]",
    ]

    for cat_name, labels in categories.items():
        preview = ", ".join(labels[:15])
        more = f" 외 {len(labels) - 15}건" if len(labels) > 15 else ""
        lines.append(f"- {cat_name} ({len(labels)}건): {preview}{more}" if labels else f"- {cat_name}: 없음")

    lines.append("")
    lines.append("[사전 산출: 차년도 제조검토대상 — 표준생약]")
    lines.append(_format_candidate_lines(manufacture.get("표준생약") or []))
    lines.append("")
    lines.append("[사전 산출: 차년도 제조검토대상 — 지표성분]")
    lines.append(_format_candidate_lines(manufacture.get("지표성분") or []))
    lines.append("")
    lines.append("[사전 산출: 모니터링 대상(분양 가속)]")
    if monitoring:
        for i, r in enumerate(monitoring, 1):
            ar = r.get("annual_rate")
            ar_txt = f"{ar:.2f}" if isinstance(ar, (int, float)) else "-"
            lines.append(
                f"{i}. {r['label']} | {r['indicators']} | "
                f"연평균:{ar_txt} | 소진:{r.get('deplete_ym') or '-'} | "
                f"{r['reliability']}"
            )
    else:
        lines.append("없음")

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
            f"5년이내고갈:{'예' if stats.get('deplete_within_5y') else '아니오'} | "
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


def build_followup_prompt(
    base_report: str,
    user_question: str,
    items: list[StockItem] | None = None,
) -> str:
    """초기 리포트 이후 추가 질문용 프롬프트."""
    lines = [
        "당신은 생약표준품 재고·분양 분석 전문가입니다.",
        "아래는 이미 생성된 초기 분석 리포트와 사용자의 후속 질문입니다.",
        "",
        "[할루시네이션 금지] 데이터·리포트에 없는 내용을 단정하지 마세요. "
        "불확실하면 '확실하지 않음' 또는 '추측입니다'라고 밝히세요.",
        "[형식] '분석 전문가의 제언' 섹션은 작성하지 마세요.",
        "",
        "[초기 분석 리포트]",
        base_report.strip() or "(리포트 없음)",
        "",
    ]
    if items:
        flags = collect_ai_analysis_flags(items)
        lines.append(
            f"[참고: 고갈후보 {len(flags.get('deplete_codes') or [])}건 / "
            f"가속 {len(flags.get('surge_codes') or [])}건 / "
            f"제조검토 표준생약 "
            f"{len((flags.get('manufacture_candidates') or {}).get('표준생약') or [])}건 · "
            f"지표성분 "
            f"{len((flags.get('manufacture_candidates') or {}).get('지표성분') or [])}건 / "
            f"환산총액 {_fmt_money((flags.get('valuation') or {}).get('total_value'))}]"
        )
        lines.append("")
    lines.append(f"[사용자 질문]\n{user_question.strip()}")
    lines.append("")
    lines.append("질문에 대해 한국어 마크다운으로 간결하고 정확하게 답변해 주세요.")
    return "\n".join(lines)


def category_counts(items: list[StockItem]) -> dict[str, int]:
    counts: dict[str, int] = defaultdict(int)
    for it in items:
        counts[it.std_type or "미분류"] += 1
    return dict(counts)

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
