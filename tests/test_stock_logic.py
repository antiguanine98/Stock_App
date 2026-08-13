"""엑셀 파싱 / 연도별 소급 보정 / AI 대상 필터 단위 테스트."""

from __future__ import annotations

import os
import sys
import tempfile
from datetime import date

import pandas as pd

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from stock_logic import (  # noqa: E402
    StockPoint,
    apply_retroactive_correction,
    build_ai_prompt,
    build_scatter3d_record,
    build_scatter3d_records,
    discover_timeseries_pairs,
    estimate_depletion,
    items_for_ai_analysis,
    load_stock_excel,
    normalize_std_type,
    scatter_cat_label,
)


def _make_sample_xlsx() -> str:
    rows = [
        {
            "순번": 1,
            "표준품구분": "생약(표준생약)",
            "관리번호": "STD-001",
            "한글명": "감초",
            "영문명": "Glycyrrhizae",
            "잔고": 120,
            "등록일자": "2020-01-01",
            "분양여부": "Y",
            "규격": "1g",
            "단위": "병",
            "가격(원)": 1000,
            "변경일자1": date(2022, 2, 1),
            "재고량1": 150,
            "변경일자2": date(2023, 3, 10),
            "재고량2": 140,
            "변경일자3": date(2024, 5, 20),
            "재고량3": 120,
        },
        {
            "순번": 2,
            "표준품구분": "생약(대조품)",
            "관리번호": "NST-014",
            "한글명": "당귀",
            "영문명": "Angelicae",
            "잔고": 8,
            "등록일자": "2020-01-02",
            "분양여부": "Y",
            "규격": "1g",
            "단위": "병",
            "가격(원)": 2000,
            "변경일자1": date(2023, 1, 1),
            "재고량1": 10,
            "변경일자2": date(2024, 6, 1),
            "재고량2": 8,
            "변경일자3": None,
            "재고량3": None,
        },
        {
            "순번": 3,
            "표준품구분": "생약(표준생약)",
            "관리번호": "STD-002",
            "한글명": "황기",
            "영문명": "Astragali",
            "잔고": 50,
            "등록일자": "2020-01-03",
            "분양여부": "N",
            "규격": "1g",
            "단위": "병",
            "가격(원)": 1500,
            "변경일자1": date(2024, 2, 10),
            "재고량1": 50,
            "변경일자2": date(2024, 3, 1),
            "재고량2": 45,
            "변경일자3": date(2024, 3, 1),
            "재고량3": 50,  # 동일일자 소급
        },
        {
            "순번": 4,
            "표준품구분": "생약(지표성분)",
            "관리번호": "STD-003",
            "한글명": "인삼",
            "영문명": "Ginseng",
            "잔고": 200,
            "등록일자": "2020-01-04",
            "분양여부": "Y",
            "규격": "1g",
            "단위": "병",
            "가격(원)": 3000,
            "변경일자1": date(2022, 1, 1),
            "재고량1": 300,
            "변경일자2": date(2023, 1, 1),
            "재고량2": 250,
            "변경일자3": date(2024, 1, 1),
            "재고량3": 200,
        },
        {
            "순번": 5,
            "표준품구분": "생약(표준생약)",
            "관리번호": "NST-020",
            "한글명": "데이터없음",
            "영문명": "Empty",
            "잔고": 0,
            "등록일자": "2020-01-05",
            "분양여부": "N",
            "규격": "1g",
            "단위": "병",
            "가격(원)": 0,
            "변경일자1": None,
            "재고량1": None,
            "변경일자2": None,
            "재고량2": None,
            "변경일자3": None,
            "재고량3": None,
        },
    ]
    df = pd.DataFrame(rows)
    fd, path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    df.to_excel(path, index=False)
    return path


def test_normalize_std_type():
    assert normalize_std_type("생약(표준생약)") == "표준생약"
    assert normalize_std_type("생약(대조품)") == "대조생약"
    assert normalize_std_type("생약(지표성분)") == "지표성분"


def test_discover_pairs_ignores_extra_columns():
    cols = [
        "순번", "표준품구분", "관리번호", "한글명", "영문명", "잔고", "등록일자", "분양여부",
        "규격", "단위", "가격(원)",
        "변경일자1", "재고량1", "변경일자2", "재고량2",
    ]
    pairs = discover_timeseries_pairs(cols)
    assert pairs == [("변경일자1", "재고량1"), ("변경일자2", "재고량2")]


def test_year_end_retroactive_correction():
    points = [
        StockPoint(date(2023, 3, 1), 40),
        StockPoint(date(2023, 8, 1), 35),
        StockPoint(date(2024, 2, 1), 50),  # 증가 → 2023 연도말 소급
        StockPoint(date(2024, 7, 1), 45),
    ]
    corrected, cnt = apply_retroactive_correction(points)
    assert [p.change_date.year for p in corrected] == [2023, 2024]
    # 2023 연도말=35, 2024 연도말=45 → 증가 10 → 2023이 45로 보정
    assert corrected[0].quantity == 45
    assert corrected[1].quantity == 45
    assert cnt == 1


def test_same_date_last_value_within_year():
    points = [
        StockPoint(date(2024, 3, 1), 45),
        StockPoint(date(2024, 2, 10), 50),
        StockPoint(date(2024, 3, 1), 50),
        StockPoint(date(2024, 7, 1), 50),
    ]
    corrected, _ = apply_retroactive_correction(points)
    assert len(corrected) == 1
    assert corrected[0].change_date == date(2024, 7, 1)
    assert corrected[0].quantity == 50


def test_load_sample_excel_and_identifiers():
    path = _make_sample_xlsx()
    try:
        df, items = load_stock_excel(path)
        assert "관리번호" in df.columns
        assert "한글명" in df.columns
        assert "재고" in df.columns
        assert "잔고" not in df.columns
        assert "등록일자" not in df.columns
        assert len(items) == 5
        assert {it.manage_no for it in items} == {
            "STD-001", "NST-014", "STD-002", "STD-003", "NST-020",
        }
        std001 = next(it for it in items if it.manage_no == "STD-001")
        assert std001.std_type == "표준생약"
        assert [p.change_date.year for p in std001.corrected_points] == [2022, 2023, 2024]
        assert [p.quantity for p in std001.corrected_points] == [150, 140, 120]

        nst = next(it for it in items if it.manage_no == "NST-014")
        assert nst.std_type == "대조생약"

        empty = next(it for it in items if it.manage_no == "NST-020")
        assert empty.corrected_points == []
        assert empty.has_stock_change is False
    finally:
        os.remove(path)


def test_ai_targets_only_changed_items():
    path = _make_sample_xlsx()
    try:
        _, items = load_stock_excel(path)
        targets = items_for_ai_analysis(items)
        codes = {it.manage_no for it in targets}
        assert "STD-001" in codes
        assert "NST-014" in codes
        assert "STD-003" in codes
        assert "NST-020" not in codes
        prompt = build_ai_prompt(targets)
        assert "민간 분양" in prompt or "생약표준품" in prompt
        assert "추가 생산" in prompt
        assert "할루시네이션" in prompt or "억측" in prompt
        assert "신뢰도" in prompt
        assert "STD-001" in prompt
        assert "NST-020" not in prompt
        stats = estimate_depletion(next(it for it in items if it.manage_no == "STD-001"))
        assert "depletion_category" in stats
        assert "reliability" in stats
        assert "priority_score" in stats
        assert stats["deplete_ym"] is None or "년" in str(stats["deplete_ym"])
    finally:
        os.remove(path)


def test_year_display_and_continuous_axis():
    from stock_logic import align_display_series, fill_continuous_years, format_year

    points = [
        StockPoint(date(2022, 5, 1), 100),
        StockPoint(date(2024, 8, 1), 80),
    ]
    filled = fill_continuous_years(points)
    assert [p.change_date.year for p in filled] == [2022, 2023, 2024]
    assert filled[1].quantity == 100  # 빈 연도 forward-fill
    assert format_year(date(2023, 1, 1)) == "2023"

    years, corr, orig = align_display_series(points, points)
    assert years == ["2022", "2023", "2024"]
    assert corr == [100, 100, 80]


def test_multi_file_merge():
    from stock_logic import process_excels

    path_a = _make_sample_xlsx()
    # 두 번째 파일: 동일 관리번호 추가 연도 + 신규 품목
    rows = [
        {
            "순번": 1,
            "표준품구분": "생약(표준생약)",
            "관리번호": "STD-001",
            "한글명": "감초",
            "영문명": "Glycyrrhizae",
            "잔고": 100,
            "등록일자": "2020-01-01",
            "분양여부": "Y",
            "변경일자1": date(2025, 1, 1),
            "재고량1": 100,
        },
        {
            "순번": 2,
            "표준품구분": "생약(지표성분)",
            "관리번호": "IDX-009",
            "한글명": "신규",
            "영문명": "New",
            "잔고": 30,
            "등록일자": "2020-01-01",
            "분양여부": "Y",
            "변경일자1": date(2023, 1, 1),
            "재고량1": 50,
            "변경일자2": date(2024, 1, 1),
            "재고량2": 30,
        },
    ]
    fd, path_b = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    pd.DataFrame(rows).to_excel(path_b, index=False)
    try:
        data = process_excels([path_a, path_b])
        codes = {it.manage_no for it in data["stock_items"]}
        assert "STD-001" in codes
        assert "IDX-009" in codes
        std001 = next(it for it in data["stock_items"] if it.manage_no == "STD-001")
        assert 2025 in {p.change_date.year for p in std001.corrected_points}
        # UI 날짜는 연도 문자열
        ui = next(it for it in data["items"] if it["mgmt_no"] == "STD-001")
        assert all(len(str(d)) == 4 and str(d).isdigit() for d in ui["dates"])
        assert data["file_name"].endswith("개 파일") or len(data["file_names"]) == 2
    finally:
        os.remove(path_a)
        os.remove(path_b)


def test_scatter_cat_label():
    assert scatter_cat_label("표준생약") == "표준생약"
    assert scatter_cat_label("대조생약") == "대조품"
    assert scatter_cat_label("", "생약(지표성분)") == "지표성분"


def test_scatter3d_records_from_sample():
    path = _make_sample_xlsx()
    try:
        _, items = load_stock_excel(path)
        std001 = next(it for it in items if it.manage_no == "STD-001")
        rec = build_scatter3d_record(std001)
        assert rec is not None
        assert rec["cat"] == "표준생약"
        assert rec["code"] == "STD-001"
        assert rec["name"] == "감초"
        assert rec["initQty"] == 150
        assert rec["balance"] == 120
        assert rec["netDrop"] == 30
        assert abs(rec["decreaseRate"] - 20.0) < 1e-6

        stats = estimate_depletion(std001)
        assert stats["annual_rate"] is not None and stats["years_left"] is not None
        assert abs(rec["annualRate"] - stats["annual_rate"]) < 1e-4
        expected_years = min(float(stats["years_left"]), 50.0)
        assert abs(rec["yearsLeft"] - expected_years) < 1e-3
        assert abs(rec["runway"] - expected_years) < 1e-3
        assert rec["yearsLeft"] <= 50.0

        empty = next(it for it in items if it.manage_no == "NST-020")
        assert build_scatter3d_record(empty) is None

        records = build_scatter3d_records(items)
        assert all("annualRate" in r and "netDrop" in r for r in records)
        assert all(r["yearsLeft"] <= 50.0 for r in records)
        # 감소 추이 있는 품목만 포함
        assert "NST-020" not in {r["code"] for r in records}
        assert next(r for r in records if r["code"] == "NST-014")["cat"] == "대조품"
    finally:
        os.remove(path)


if __name__ == "__main__":
    import traceback

    tests = [
        test_normalize_std_type,
        test_discover_pairs_ignores_extra_columns,
        test_year_end_retroactive_correction,
        test_same_date_last_value_within_year,
        test_load_sample_excel_and_identifiers,
        test_ai_targets_only_changed_items,
        test_year_display_and_continuous_axis,
        test_multi_file_merge,
        test_scatter_cat_label,
        test_scatter3d_records_from_sample,
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
