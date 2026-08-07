"""재고 엑셀 파싱 · 소급 보정 · AI 프롬프트 구성 로직."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Any, Optional

import pandas as pd

BASE_COLUMNS = [
    "순번",
    "표준품구분",
    "관리번호",
    "한글명",
    "영문명",
    "잔고",
    "등록일자",
    "분양여부",
]
BASE_COL_COUNT = len(BASE_COLUMNS)  # A~H (I열 이전)


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
    """변경일자를 YYYY-MM-DD date 로 파싱. 빈값/NaN 은 None."""
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
    try:
        ts = pd.to_datetime(text, errors="coerce")
        if pd.isna(ts):
            return None
        return ts.date()
    except Exception:
        return None


def parse_qty(value: Any) -> Optional[float]:
    if _is_empty(value):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def format_date(d: date) -> str:
    return d.strftime("%Y-%m-%d")


@dataclass
class StockPoint:
    change_date: date
    quantity: float


@dataclass
class StockItem:
    seq: Any = None
    std_type: Any = None
    manage_no: str = ""
    name_ko: str = ""
    name_en: Any = None
    balance: Any = None
    registered_at: Any = None
    distributed: Any = None
    raw_points: list[StockPoint] = field(default_factory=list)
    corrected_points: list[StockPoint] = field(default_factory=list)

    @property
    def label(self) -> str:
        return f"{self.manage_no}({self.name_ko})"

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
        """최초 대비 최종 수량 변화가 있는 품목만 True."""
        delta = self.qty_delta
        if delta is None:
            return False
        return abs(delta) > 1e-9


def apply_retroactive_correction(points: list[StockPoint]) -> list[StockPoint]:
    """동일 변경일자에 여러 재고량이 있으면 마지막 값으로 소급 보정.

    엑셀 열 순서상 뒤에 나온 쌍이 보정 기록으로 간주되며,
    날짜별 최종 수치만 남긴 뒤 날짜 오름차순으로 정렬한다.
    """
    if not points:
        return []
    by_date: dict[date, float] = {}
    for point in points:
        by_date[point.change_date] = point.quantity
    return [
        StockPoint(change_date=d, quantity=by_date[d])
        for d in sorted(by_date.keys())
    ]


def extract_change_pairs(row: pd.Series, columns: list[str]) -> list[StockPoint]:
    """I열(인덱스 8) 이후 변경일자/재고량 쌍을 추출. NaN·빈값 제외."""
    points: list[StockPoint] = []
    pair_cols = columns[BASE_COL_COUNT:]
    i = 0
    while i < len(pair_cols):
        date_col = pair_cols[i]
        qty_col = pair_cols[i + 1] if i + 1 < len(pair_cols) else None
        if qty_col is None:
            break
        d = parse_date(row.get(date_col))
        q = parse_qty(row.get(qty_col))
        if d is not None and q is not None:
            points.append(StockPoint(change_date=d, quantity=q))
        i += 2
    return points


def _cell_str(value: Any) -> str:
    if _is_empty(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def build_corrected_dataframe(items: list[StockItem]) -> pd.DataFrame:
    """소급 보정 최종 수치만 반영한 표용 DataFrame 생성."""
    max_points = max((len(it.corrected_points) for it in items), default=0)
    records: list[dict[str, Any]] = []
    for it in items:
        reg = parse_date(it.registered_at)
        rec: dict[str, Any] = {
            "순번": it.seq,
            "표준품구분": it.std_type,
            "관리번호": it.manage_no,
            "한글명": it.name_ko,
            "영문명": it.name_en,
            "잔고": it.balance,
            "등록일자": format_date(reg) if reg else it.registered_at,
            "분양여부": it.distributed,
        }
        for idx in range(max_points):
            if idx < len(it.corrected_points):
                pt = it.corrected_points[idx]
                rec[f"변경일자{idx + 1}"] = format_date(pt.change_date)
                rec[f"재고량{idx + 1}"] = pt.quantity
            else:
                rec[f"변경일자{idx + 1}"] = None
                rec[f"재고량{idx + 1}"] = None
        records.append(rec)

    columns = list(BASE_COLUMNS)
    for idx in range(max_points):
        columns += [f"변경일자{idx + 1}", f"재고량{idx + 1}"]
    return pd.DataFrame(records, columns=columns)


def load_stock_excel(path: str) -> tuple[pd.DataFrame, list[StockItem]]:
    """엑셀을 읽어 소급 보정된 StockItem 목록과 보정 데이터표용 DataFrame을 반환."""
    df = pd.read_excel(path, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in ("관리번호", "한글명") if c not in df.columns]
    if missing:
        raise ValueError(f"필수 컬럼이 없습니다: {', '.join(missing)}")

    for col in BASE_COLUMNS:
        if col not in df.columns:
            df[col] = None

    # 기본정보 컬럼을 앞으로, 이후 열은 변동기록 쌍으로 유지
    other_cols = [c for c in df.columns if c not in BASE_COLUMNS]
    df = df[BASE_COLUMNS + other_cols]

    items: list[StockItem] = []
    for _, row in df.iterrows():
        manage_no = _cell_str(row.get("관리번호"))
        name_ko = _cell_str(row.get("한글명"))
        if not manage_no and not name_ko:
            continue

        raw_points = extract_change_pairs(row, list(df.columns))
        corrected = apply_retroactive_correction(raw_points)
        items.append(
            StockItem(
                seq=row.get("순번"),
                std_type=row.get("표준품구분"),
                manage_no=manage_no,
                name_ko=name_ko,
                name_en=row.get("영문명"),
                balance=row.get("잔고"),
                registered_at=row.get("등록일자"),
                distributed=row.get("분양여부"),
                raw_points=raw_points,
                corrected_points=corrected,
            )
        )

    return build_corrected_dataframe(items), items


def items_for_ai_analysis(items: list[StockItem]) -> list[StockItem]:
    """수량 변화가 0이거나 변동 데이터가 없는 품목 제외."""
    return [it for it in items if it.has_stock_change]


def build_ai_prompt(items: list[StockItem]) -> str:
    lines = [
        "당신은 재고 데이터 분석 전문가입니다.",
        "아래는 소급 보정이 완료된 최종 재고 변동 데이터입니다.",
        "최초 수량 대비 최종 수량이 변화한 품목만 포함되어 있습니다.",
        "품목별 증감 패턴, 급격한 변동, 공통 리스크/시사점을 한국어로 간결히 분석해 주세요.",
        "",
        "[재고 변동 핵심 데이터]",
    ]
    for it in items:
        history = ", ".join(
            f"{format_date(p.change_date)}={p.quantity:g}" for p in it.corrected_points
        )
        lines.append(
            f"- 관리번호: {it.manage_no} / 한글명: {it.name_ko} / "
            f"표준품구분: {it.std_type} / 분양여부: {it.distributed} / "
            f"최초: {it.first_qty:g} → 최종: {it.last_qty:g} "
            f"(변화량: {it.qty_delta:+g}) / 추이: [{history}]"
        )
    lines.append("")
    lines.append("분석 리포트를 작성해 주세요.")
    return "\n".join(lines)
