"""엑셀 파싱 / 소급 보정 / AI 대상 필터 단위 테스트."""

from __future__ import annotations

import os
import sys
from datetime import date

import pandas as pd

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from stock_logic import (  # noqa: E402
    StockPoint,
    apply_retroactive_correction,
    build_ai_prompt,
    items_for_ai_analysis,
    load_stock_excel,
)


SAMPLE = os.path.join(ROOT, "sample", "테스트파일1.xlsx")


def test_retroactive_correction_keeps_last_same_date():
    points = [
        StockPoint(date(2024, 3, 1), 45),
        StockPoint(date(2024, 2, 10), 50),
        StockPoint(date(2024, 3, 1), 50),  # 소급 보정
        StockPoint(date(2024, 7, 1), 50),
    ]
    corrected = apply_retroactive_correction(points)
    assert [p.change_date for p in corrected] == [
        date(2024, 2, 10),
        date(2024, 3, 1),
        date(2024, 7, 1),
    ]
    assert corrected[1].quantity == 50


def test_load_sample_excel_and_identifiers():
    df, items = load_stock_excel(SAMPLE)
    assert "관리번호" in df.columns
    assert "한글명" in df.columns
    assert len(items) == 5
    assert {it.manage_no for it in items} == {
        "STD-001",
        "NST-014",
        "STD-002",
        "STD-003",
        "NST-020",
    }
    # STD-002: 동일일자 보정으로 최종 추이는 50→50→50, 변화 없음
    std002 = next(it for it in items if it.manage_no == "STD-002")
    assert len(std002.corrected_points) == 3
    assert std002.corrected_points[1].quantity == 50
    assert std002.has_stock_change is False


def test_chart_points_exclude_empty_and_use_dates():
    _, items = load_stock_excel(SAMPLE)
    std001 = next(it for it in items if it.manage_no == "STD-001")
    assert [p.change_date.isoformat() for p in std001.corrected_points] == [
        "2024-02-01",
        "2024-03-10",
        "2024-05-20",
    ]
    assert [p.quantity for p in std001.corrected_points] == [150, 140, 120]

    empty = next(it for it in items if it.manage_no == "NST-020")
    assert empty.corrected_points == []
    assert empty.has_stock_change is False


def test_ai_targets_only_changed_items():
    _, items = load_stock_excel(SAMPLE)
    targets = items_for_ai_analysis(items)
    codes = {it.manage_no for it in targets}
    # 변화 있음: STD-001(150→120), NST-014(10→8), STD-003(300→200)
    # 제외: STD-002(변화0), NST-020(데이터없음)
    assert codes == {"STD-001", "NST-014", "STD-003"}
    prompt = build_ai_prompt(targets)
    assert "STD-001" in prompt
    assert "STD-002" not in prompt
    assert "NST-020" not in prompt


def test_corrected_dataframe_pair_columns():
    df, _ = load_stock_excel(SAMPLE)
    assert list(df.columns[:8]) == [
        "순번",
        "표준품구분",
        "관리번호",
        "한글명",
        "영문명",
        "잔고",
        "등록일자",
        "분양여부",
    ]
    # I열 이후 쌍
    assert df.columns[8] == "변경일자1"
    assert df.columns[9] == "재고량1"
    # NaN 아닌 날짜는 YYYY-MM-DD 문자열
    val = df.loc[df["관리번호"] == "STD-001", "변경일자1"].iloc[0]
    assert val == "2024-02-01"


if __name__ == "__main__":
    import traceback

    tests = [
        test_retroactive_correction_keeps_last_same_date,
        test_load_sample_excel_and_identifiers,
        test_chart_points_exclude_empty_and_use_dates,
        test_ai_targets_only_changed_items,
        test_corrected_dataframe_pair_columns,
    ]
    failed = 0
    for fn in tests:
        try:
            fn()
            print(f"PASS {fn.__name__}")
        except Exception:
            failed += 1
            print(f"FAIL {fn.__name__}")
            traceback.print_exc()
    raise SystemExit(failed)
