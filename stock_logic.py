"""생약표준품 재고 엑셀 파싱 · 연도별 소급 보정 · AI 프롬프트 구성."""

from __future__ import annotations

import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Optional

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
    return d.strftime("%Y-%m-%d")


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
    distributed: Any = None
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


def build_corrected_dataframe(
    items: list[StockItem],
    meta_cols: list[str],
) -> pd.DataFrame:
    """소급 보정 최종 수치만 반영한 표용 DataFrame (등록일자 제외, 잔고→재고)."""
    max_points = max((len(it.corrected_points) for it in items), default=0)
    display_meta = [display_column_name(c) for c in meta_cols]
    records: list[dict[str, Any]] = []

    for it in items:
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
            if idx < len(it.corrected_points):
                pt = it.corrected_points[idx]
                rec[f"변경일자{idx + 1}"] = format_date(pt.change_date)
                rec[f"재고량{idx + 1}"] = pt.quantity
            else:
                rec[f"변경일자{idx + 1}"] = None
                rec[f"재고량{idx + 1}"] = None
        records.append(rec)

    columns = list(display_meta)
    for idx in range(max_points):
        columns += [f"변경일자{idx + 1}", f"재고량{idx + 1}"]
    return pd.DataFrame(records, columns=columns)


def load_stock_excel(path: str) -> tuple[pd.DataFrame, list[StockItem]]:
    """엑셀을 읽어 연도별 소급 보정된 StockItem 목록과 표용 DataFrame을 반환."""
    df = pd.read_excel(path, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in ("관리번호", "한글명") if c not in df.columns]
    if missing:
        raise ValueError(f"필수 컬럼이 없습니다: {', '.join(missing)}")

    meta_cols, pairs = classify_columns(list(df.columns))
    if not pairs:
        raise ValueError(
            "'변경일자'/'재고량' 키워드가 포함된 시계열 컬럼 쌍을 찾을 수 없습니다."
        )

    known = {"순번", "표준품구분", "관리번호", "한글명", "영문명", "재고", "잔고", "분양여부"}
    items: list[StockItem] = []

    for _, row in df.iterrows():
        manage_no = _cell_str(row.get("관리번호"))
        name_ko = _cell_str(row.get("한글명"))
        if not manage_no and not name_ko:
            continue

        raw_points = extract_change_pairs(row, pairs)
        year_end = collapse_to_year_end(raw_points)
        corrected, corr_cnt = apply_retroactive_correction(raw_points)

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
                distributed=row.get("분양여부"),
                extra_meta=extra,
                raw_points=raw_points,
                year_end_points=year_end,
                corrected_points=corrected,
                correction_count=corr_cnt,
            )
        )

    return build_corrected_dataframe(items, meta_cols), items


def process_excel(file_path: str) -> dict[str, Any]:
    """GUI용 통합 처리 결과."""
    table_df, items = load_stock_excel(file_path)
    meta_cols = [c for c in table_df.columns if not str(c).startswith(("변경일자", "재고량"))]
    max_pairs = max((len(it.corrected_points) for it in items), default=0)
    total_corrections = sum(it.correction_count for it in items)
    categories = sorted({it.std_type for it in items if it.std_type})

    ui_items: list[dict[str, Any]] = []
    for idx, it in enumerate(items):
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
                meta_vals.append(_cell_str(it.extra_meta.get(col, table_df.iloc[idx].get(col, ""))))

        year_dates = [format_date(p.change_date) for p in it.year_end_points]
        year_orig = [p.quantity for p in it.year_end_points]
        corr_dates = [format_date(p.change_date) for p in it.corrected_points]
        corr_qtys = [p.quantity for p in it.corrected_points]

        ui_items.append(
            {
                "row_index": idx,
                "meta": meta_vals,
                "mgmt_no": it.manage_no,
                "korean_name": it.name_ko,
                "std_type": it.std_type,
                "label": it.label,
                "dates": corr_dates,
                "original_dates": year_dates,
                "original": year_orig,
                "corrected": corr_qtys,
                "has_change": it.has_stock_change,
                "delta": it.qty_delta if it.qty_delta is not None else 0.0,
                "first_qty": it.first_qty,
                "last_qty": it.last_qty,
                "correction_cells": it.correction_count,
                "stock_item": it,
            }
        )

    return {
        "file_path": file_path,
        "file_name": Path(file_path).name,
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


def estimate_depletion(item: StockItem) -> dict[str, Any]:
    """분양 속도·소진 예상 기간(대략) 산출."""
    pts = item.corrected_points
    if len(pts) < 2 or item.first_qty is None or item.last_qty is None:
        return {
            "speed": "데이터부족",
            "annual_rate": None,
            "years_left": None,
            "deplete_within_5y": False,
        }

    days = (pts[-1].change_date - pts[0].change_date).days or 1
    years = max(days / 365.25, 1 / 365.25)
    net_drop = item.first_qty - item.last_qty
    annual_rate = net_drop / years  # 양수면 감소 속도

    if annual_rate <= 1e-9:
        speed = "느림" if net_drop <= 0 else "보통"
        years_left = None
        deplete = False
    else:
        years_left = item.last_qty / annual_rate if item.last_qty > 0 else 0.0
        if annual_rate >= 40:
            speed = "빠름"
        elif annual_rate >= 10:
            speed = "보통"
        else:
            speed = "느림"
        deplete = years_left is not None and years_left <= 5.0

    # 최근 구간 급증(감소 가속) 여부
    recent_surge = False
    if len(pts) >= 3:
        mid = len(pts) // 2
        early = pts[0].quantity - pts[mid].quantity
        late = pts[mid].quantity - pts[-1].quantity
        early_years = max((pts[mid].change_date - pts[0].change_date).days / 365.25, 1e-6)
        late_years = max((pts[-1].change_date - pts[mid].change_date).days / 365.25, 1e-6)
        if late / late_years > (early / early_years) * 2 and late > 0:
            recent_surge = True

    return {
        "speed": speed,
        "annual_rate": annual_rate,
        "years_left": years_left,
        "deplete_within_5y": deplete,
        "recent_surge": recent_surge,
    }


def build_ai_prompt(items: list[StockItem]) -> str:
    targets = items_for_ai_analysis(items)
    bulk_dates = detect_bulk_decrease_dates(items)

    lines = [
        "당신은 생약표준품 재고·분양 분석 전문가입니다.",
        "아래는 연도별 최종 재고량 기준으로 소급 보정이 완료된 데이터입니다.",
        "",
        "[기본 전제]",
        "1. 생약표준품은 민간 분양에 따라 지속적으로 감소하는 것이 정상이므로, 일반적인 감소 추이는 정상으로 간주합니다.",
        "2. 재고량이 증가하도록 반영된 소급 보정은 주기적 전수조사 결과에 따른 조정이므로 오류가 아닙니다.",
        "3. 동일 날짜에 100개 이상 품목이 동시 감소한 기록은 연구 과제 목적의 대량 출고이므로 민간 분양 분석에서 제외합니다.",
        "",
        "[분석 요청 항목]",
        "1. 품목별 분양 속도(빠름/보통/느림) 추이 분석",
        "2. 현재 분양 속도 기준 재고 소진 예상 기간(남은 기간) 예측",
        "3. 5년 이내 재고 소진(고갈)이 예상되는 품목은 반드시 '추가 생산(제조) 필요' 경고와 목록으로 보고",
        "4. 최근 분양량이 갑자기 급증한 특이 품목 선별 및 원인 분석",
        "",
        f"[연구과제 대량출고로 제외할 날짜] {', '.join(bulk_dates) if bulk_dates else '해당 없음'}",
        "",
        "[재고 변동 핵심 데이터]",
    ]

    deplete_list: list[str] = []
    surge_list: list[str] = []

    for it in targets:
        stats = estimate_depletion(it)
        history = ", ".join(
            f"{format_date(p.change_date)}={p.quantity:g}" for p in it.corrected_points
        )
        years_left = stats["years_left"]
        years_txt = f"{years_left:.1f}년" if years_left is not None else "산출불가/감소없음"
        lines.append(
            f"- {it.label} | 구분:{it.std_type} | 분양여부:{_cell_str(it.distributed)} | "
            f"최초:{it.first_qty:g} → 최종:{it.last_qty:g} (Δ{it.qty_delta:+g}) | "
            f"분양속도:{stats['speed']} | 예상소진:{years_txt} | "
            f"5년이내고갈:{'예' if stats['deplete_within_5y'] else '아니오'} | "
            f"최근급증:{'예' if stats['recent_surge'] else '아니오'} | 추이:[{history}]"
        )
        if stats["deplete_within_5y"]:
            deplete_list.append(it.label)
        if stats["recent_surge"]:
            surge_list.append(it.label)

    lines.append("")
    lines.append(
        f"[사전 산출: 5년 이내 고갈 후보] {', '.join(deplete_list) if deplete_list else '없음'}"
    )
    lines.append(
        f"[사전 산출: 최근 분양 급증 후보] {', '.join(surge_list) if surge_list else '없음'}"
    )
    lines.append("")
    lines.append("위 전제와 요청 항목에 맞는 한국어 마크다운 분석 리포트를 작성해 주세요.")
    return "\n".join(lines)


def category_counts(items: list[StockItem]) -> dict[str, int]:
    counts: dict[str, int] = defaultdict(int)
    for it in items:
        counts[it.std_type or "미분류"] += 1
    return dict(counts)
