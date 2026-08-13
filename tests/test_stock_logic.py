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
        assert "추가 생산" in prompt or "제조검토" in prompt
        assert "할루시네이션" in prompt or "억측" in prompt
        assert "신뢰도" in prompt
        assert "1페이지 요약 대시보드" in prompt or "핵심 KPI" in prompt
        assert "분석 전문가의 제언" in prompt  # 금지 지시문에 포함
        assert "제언' 섹션은 작성하지 마세요" in prompt or "제언 섹션은" in prompt
        assert "STD-001" in prompt
        assert "NST-020" not in prompt
        stats = estimate_depletion(next(it for it in items if it.manage_no == "STD-001"))
        assert "depletion_category" in stats
        assert "reliability" in stats
        assert stats["reliability"].get("grade") in ("A", "B", "C", "D")
        assert "priority_score" in stats
        assert "acceleration" in stats
        assert stats["deplete_ym"] is None or "년" in str(stats["deplete_ym"])
        # 가격 환산
        assert stats.get("unit_price") == 1000
        assert stats.get("stock_value") == 120 * 1000
        from stock_logic import collect_ai_analysis_flags, compute_inventory_valuation

        flags = collect_ai_analysis_flags(items)
        assert "dashboard" in flags
        assert "valuation" in flags
        assert flags["valuation"]["total_value"] > 0
        assert len(flags["dashboard"]["kpis"]) >= 8
        assert len(flags["dashboard"]["summary_lines"]) >= 4
        val = compute_inventory_valuation(items)
        assert any(r["manage_no"] == "STD-001" for r in val["top20"])
    finally:
        os.remove(path)


def test_decrease_only_excludes_increases():
    from stock_logic import StockItem, StockPoint, decrease_only_rate_stats, estimate_depletion

    # 감소 후 증가(전수조사) 후 감소
    points = [
        StockPoint(date(2020, 1, 1), 100),
        StockPoint(date(2021, 1, 1), 80),   # -20
        StockPoint(date(2022, 1, 1), 100),  # +20 증가 → 제외
        StockPoint(date(2023, 1, 1), 70),   # -30
    ]
    dec = decrease_only_rate_stats(points)
    assert dec["increase_segments"] == 1
    assert dec["decrease_segments"] == 2
    assert dec["annual_rate"] is not None
    assert abs(dec["total_drop"] - 50) < 1e-6

    item = StockItem(
        manage_no="T-1",
        name_ko="테스트",
        std_type="표준생약",
        unit_price=500,
        raw_points=list(points),
        year_end_points=list(points),
        corrected_points=list(points),
    )
    stats = estimate_depletion(item)
    assert stats["increase_segments_excluded"] >= 1
    assert stats["acceleration"] in ("급가속", "증가", "안정", "감소")
    assert stats["stock_value"] == 70 * 500


def test_live_query_followup_uses_inventory():
    from stock_logic import (
        StockItem,
        StockPoint,
        build_followup_prompt,
        query_live_inventory_context,
    )

    item = StockItem(
        manage_no="STD-001",
        name_ko="감초",
        std_type="표준생약",
        unit_price=1000,
        raw_points=[
            StockPoint(date(2022, 1, 1), 150),
            StockPoint(date(2023, 1, 1), 140),
            StockPoint(date(2024, 1, 1), 120),
        ],
        year_end_points=[
            StockPoint(date(2022, 1, 1), 150),
            StockPoint(date(2023, 1, 1), 140),
            StockPoint(date(2024, 1, 1), 120),
        ],
        corrected_points=[
            StockPoint(date(2022, 1, 1), 150),
            StockPoint(date(2023, 1, 1), 140),
            StockPoint(date(2024, 1, 1), 120),
        ],
    )
    live = query_live_inventory_context([item], "감초 소진 언제야?")
    assert "실시간 재고" in live
    assert "STD-001" in live or "감초" in live
    prompt = build_followup_prompt(
        "초기 리포트 요약",
        "감초 소진 예상 시점은?",
        [item],
        compendium_context="[공정서 DB]\n- 감초 | 기원=콩과",
    )
    assert "실시간 재고" in prompt
    assert "공정서" in prompt
    assert "초기 리포트보다 이 수치를 우선" in prompt or "실시간 재검토" in prompt


def test_compendium_db_not_timeseries():
    from stock_logic import (
        StockItem,
        format_compendium_context,
        load_compendium_excel,
        build_ai_prompt,
    )

    rows = [
        {
            "생약명": "감초",
            "관리번호": "STD-001",
            "기원": "콩과",
            "성상": "황색",
            "확인시험": "TLC",
        },
        {
            "생약명": "당귀",
            "관리번호": "NST-014",
            "기원": "미나리과",
            "성상": "갈색",
            "확인시험": "HPLC",
        },
    ]
    fd, path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    pd.DataFrame(rows).to_excel(path, index=False)
    try:
        data = load_compendium_excel(path)
        assert data["row_count"] == 2
        assert "dataframe" in data
        assert "변경일자" not in data["columns"]
        items = [
            StockItem(manage_no="STD-001", name_ko="감초", std_type="표준생약"),
        ]
        ctx = format_compendium_context(data["dataframe"], items, meta=data)
        assert "공정서" in ctx
        assert "감초" in ctx
        assert "규격" in ctx or "참조" in ctx
        prompt = build_ai_prompt([], compendium_context=ctx)
        assert "공정서" in prompt
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
        test_decrease_only_excludes_increases,
        test_live_query_followup_uses_inventory,
        test_compendium_db_not_timeseries,
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
