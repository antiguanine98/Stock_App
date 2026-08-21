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
    StockItem,
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
        collect_ai_analysis_flags,
        find_items_by_partial_query,
        markdown_report_to_collapsible_html,
        query_live_inventory_context,
        serialize_flags_snapshot,
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
    tanshinone = StockItem(
        manage_no="STD-099",
        name_ko="탄시논 IIA",
        std_type="지표성분",
        unit_price=500,
        corrected_points=[StockPoint(date(2024, 1, 1), 10)],
    )
    partial = find_items_by_partial_query([item, tanshinone], "탄시논 소진?")
    assert any(it.name_ko == "탄시논 IIA" for it in partial)

    flags = collect_ai_analysis_flags([item])
    live = query_live_inventory_context([item], "감초 소진 언제야?", flags=flags)
    assert "실시간 재고" in live
    assert "STD-001" in live or "감초" in live
    assert "스냅샷" in live or "by_code" in live
    snap = serialize_flags_snapshot(flags)
    assert "재고 조사 스냅샷" in snap or "스냅샷" in snap
    assert "전수" in snap
    assert "위험등급" in snap or "소진구간" in snap
    prompt = build_followup_prompt(
        "초기 리포트 요약 — 가스트로디게닌, 백출 등 전량 81건",
        "모니터링 대상 81건 모두 알려줘",
        [item],
        compendium_context="[공정서 DB]\n- 감초 | 기원=콩과",
        flags=flags,
    )
    assert "실시간 재고" in prompt
    assert "공정서" in prompt
    assert "스냅샷" in prompt
    assert "전수 목록" in prompt
    assert "초기 리포트보다 이 수치를 우선" in prompt or "실시간 재검토" in prompt

    from stock_logic import detect_full_list_intent

    intent = detect_full_list_intent("모니터링 대상 81건 모두 알려줘")
    assert intent["wants_full"] and intent["monitoring"]
    intent2 = detect_full_list_intent("급가속 품목 전체 리스트")
    assert intent2["wants_full"] and intent2["monitoring"]

    html = markdown_report_to_collapsible_html(
        "- 1년 이내 (10건): "
        + ", ".join(f"품목{i}" for i in range(12))
    )
    # 접기/토글 배제 — 전수 표시
    assert "expand:" not in html
    assert "collapse:" not in html
    assert "details" not in html
    assert "품목0" in html and "품목11" in html


def test_full_catalog_snapshot_and_intent():
    from stock_logic import (
        StockItem,
        StockPoint,
        collect_ai_analysis_flags,
        detect_full_list_intent,
        query_live_inventory_context,
        serialize_flags_snapshot,
    )

    # 과거 느린 감소 → 최근 빠른 감소 = 급가속
    items = []
    for i in range(5):
        items.append(
            StockItem(
                manage_no=f"SURGE-{i:03d}",
                name_ko=f"급가속품{i}",
                std_type="표준생약",
                unit_price=100,
                corrected_points=[
                    StockPoint(date(2018, 1, 1), 200),
                    StockPoint(date(2019, 1, 1), 198),
                    StockPoint(date(2020, 1, 1), 196),
                    StockPoint(date(2021, 1, 1), 194),
                    StockPoint(date(2022, 1, 1), 150),
                    StockPoint(date(2023, 1, 1), 100),
                    StockPoint(date(2024, 1, 1), 40),
                ],
            )
        )

    flags = collect_ai_analysis_flags(items)
    assert "monitoring_targets" in flags
    assert "depletion_category_items" in flags
    assert "risk_grade_items" in flags
    mon = flags["monitoring_targets"]
    assert len(mon) >= 1
    assert all("name_ko" in r and "manage_no" in r for r in mon)
    for grade in ("위험", "경계", "주의", "안정"):
        assert grade in flags["risk_grade_items"]

    snap = serialize_flags_snapshot(flags)
    assert "전수" in snap and "생략 없음" in snap
    for r in mon:
        assert (r.get("name_ko") or r.get("label")) in snap
        assert str(r.get("manage_no")) in snap
    # 더 이상 상위 15건만 자르지 않음
    assert "monitoring_targets" not in snap or "전수" in snap

    intent = detect_full_list_intent("모니터링 대상 전체 알려줘")
    assert intent["wants_full"] and intent["monitoring"]
    live = query_live_inventory_context(
        items, "모니터링 대상 5건 모두 보여줘", flags=flags
    )
    assert "실시간 전수: 모니터링" in live
    for r in mon:
        assert str(r.get("manage_no")) in live


def test_split_markdown_report_sections():
    from stock_logic import split_markdown_report_sections, _report_section_short_label

    md = (
        "## 1페이지 요약 대시보드 (핵심 KPI)\n\n| a | b |\n| --- | --- |\n| 1 | 2 |\n\n"
        "## 소진 예상 기간\n\n내용A\n\n"
        "### 재고 없음(미보유)\n\n제로품목\n\n"
        "## 차년도 제조검토대상\n\n내용B\n\n"
        "## 공정서 DB 매칭 및 수재 현황\n\n"
        "### 공정서 미보유 표준품 전수\n\n지황\n"
    )
    secs = split_markdown_report_sections(md)
    by_id = {s["id"]: s for s in secs}
    assert by_id["summary"]["short"] == "요약"
    assert by_id["deplete"]["short"] == "소진"
    assert by_id["missing"]["short"] == "미보유"
    assert "제로품목" in by_id["missing"]["markdown"] or "지황" in by_id["missing"]["markdown"]
    assert by_id["manufacture"]["short"] == "검토"
    assert "내용B" in by_id["manufacture"]["markdown"]
    assert by_id["compendium"]["short"] == "공정서"
    # 고정 버튼용 섹션은 본문 없어도 항상 존재
    assert by_id["accel"]["short"] == "가속"
    assert _report_section_short_label("모니터링 대상") == "가속"
    assert _report_section_short_label("재고 없음(미보유)") == "미보유"
    assert _report_section_short_label("차년도 제조검토대상") == "검토"


def test_markdown_report_renders_tables_as_html():
    from stock_logic import (
        format_kpi_dashboard_markdown,
        markdown_report_to_collapsible_html,
    )

    md = format_kpi_dashboard_markdown(
        {
            "kpis": [
                {"label": "대상품목 수", "display": "491종"},
                {"label": "1년 내 소진예상", "display": "10종"},
                {"label": "1~3년 소진예상", "display": "20종"},
                {"label": "3~5년 소진예상", "display": "14종"},
            ],
            "summary_lines": ["요약 의견"],
        }
    )
    html = markdown_report_to_collapsible_html(md)
    assert "<table" in html
    assert "<th" in html and "<td" in html
    assert "491종" in html and "대상품목 수" in html
    assert "| 지표 |" not in html
    assert "| --- |" not in html


def test_manufacture_candidates_always_top10_by_score():
    """5년 이내 소진이 없어도 유형별 우선순위 상위 10건을 채운다."""
    from stock_logic import select_manufacture_candidates

    items: list[StockItem] = []
    for i in range(12):
        items.append(
            StockItem(
                manage_no=f"S{i:02d}",
                name_ko=f"표준{i}",
                std_type="표준생약",
                corrected_points=[
                    StockPoint(date(2020, 1, 1), 500),
                    StockPoint(date(2021, 1, 1), 495),
                    StockPoint(date(2022, 1, 1), 490),
                    StockPoint(date(2023, 1, 1), 485),
                    StockPoint(date(2024, 1, 1), 480 - i),
                ],
            )
        )
    for i in range(3):
        items.append(
            StockItem(
                manage_no=f"M{i:02d}",
                name_ko=f"지표{i}",
                std_type="지표성분",
                corrected_points=[
                    StockPoint(date(2022, 1, 1), 100),
                    StockPoint(date(2023, 1, 1), 80),
                    StockPoint(date(2024, 1, 1), 50 - i),
                ],
            )
        )
    # 대조생약은 제외
    items.append(
        StockItem(
            manage_no="C01",
            name_ko="대조",
            std_type="대조생약",
            corrected_points=[
                StockPoint(date(2023, 1, 1), 10),
                StockPoint(date(2024, 1, 1), 1),
            ],
        )
    )

    result = select_manufacture_candidates(items)
    assert len(result["표준생약"]) == 10
    assert len(result["지표성분"]) == 3
    scores = [r["priority_score"] for r in result["표준생약"]]
    assert scores == sorted(scores, reverse=True)
    assert all(not r.get("deplete_within_5y") for r in result["표준생약"])
    assert "C01" not in {r["manage_no"] for r in result["표준생약"]}


def test_compendium_missing_set_and_followup_filter():
    from stock_logic import (
        CompendiumEntry,
        StockItem,
        attach_compendium_match_to_flags,
        detect_full_list_intent,
        filter_missing_compendium_items,
        format_compendium_stats_markdown,
        match_compendium_inventory,
        serialize_flags_snapshot,
        _norm_key,
    )

    assert _norm_key("감 초 (A)") == _norm_key("감초A")
    entries = [
        CompendiumEntry(name_ko="감초", pharmacopoeia="KP"),
        CompendiumEntry(name_ko="지황", pharmacopoeia="KHP"),
        CompendiumEntry(name_ko="백출 (규격)", pharmacopoeia="생약규격집"),
        CompendiumEntry(name_ko="당귀", pharmacopoeia="KP"),
    ]
    items = [StockItem(manage_no="1", name_ko="감초", std_type="표준생약")]
    match = match_compendium_inventory(entries, items)
    assert match["stats"]["missing_count"] == 3
    assert match["stats"]["inventory_matched"] == 1
    assert len(match["missing_items"]) == 3
    md = format_compendium_stats_markdown(match)
    assert "공정서 DB 매칭" in md and "3건" in md
    assert "미보유" in md and "|" in md  # 마크다운 표 포함

    intent = detect_full_list_intent("공정서 미보유 품목 전체 알려줘")
    assert intent["missing_compendium"]
    khp = filter_missing_compendium_items(
        match["missing_items"], "그 중 KHP 수재 품목만 골라줘"
    )
    assert len(khp) == 2
    assert all(r["pharmacopoeia_kind"] == "KHP" for r in khp)
    jih = filter_missing_compendium_items(
        match["missing_items"], "미보유 중 지황 관련 품목 있어?"
    )
    assert len(jih) == 1 and jih[0]["name_ko"] == "지황"

    flags = attach_compendium_match_to_flags({"by_code": {}}, match)
    snap = serialize_flags_snapshot(flags)
    assert "공정서 미보유" in snap
    assert "지황" in snap and "당귀" in snap


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
        test_full_catalog_snapshot_and_intent,
        test_markdown_report_renders_tables_as_html,
        test_split_markdown_report_sections,
        test_manufacture_candidates_always_top10_by_score,
        test_compendium_missing_set_and_followup_filter,
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
