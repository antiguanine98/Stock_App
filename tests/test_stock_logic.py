"""엑셀 파싱 / 연도별 소급 보정 / AI 대상 필터 단위 테스트."""

from __future__ import annotations

import os
import sys
import tempfile
from datetime import date
from pathlib import Path

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
    collapse_to_year_end,
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
        assert "제조" in prompt or "우선" in prompt or "Token Diet" in prompt
        assert "할루시네이션" in prompt or "지어내지" in prompt or "임의로" in prompt
        assert "신뢰도" in prompt or "Token Diet" in prompt or "KPI" in prompt
        assert "1페이지 요약 대시보드" in prompt or "핵심 KPI" in prompt
        assert "제언" in prompt  # 금지 지시문에 포함
        assert "STD-001" in prompt or "Token Diet" in prompt
        assert "NST-020" not in prompt or "Token Diet" in prompt
        assert "추이:[" not in prompt
        assert len(prompt) < 12000
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
    assert stats["acceleration"] in (
        "신규수요",
        "급증",
        "증가",
        "안정",
        "감소",
        "급감",
        "분석불가/자료부족",
    )
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
        "## 종합현황 대시보드\n\n| a | b |\n| --- | --- |\n| 1 | 2 |\n\n"
        "## 제조 시급 품목 리스트\n\n내용A\n\n"
        "## 과다재고 및 제조 축소 권고 품목\n\n축소내용\n\n"
        "## 공정서 미확보 표준생약 개발 대상\n\n지황\n\n"
        "## 로드맵 총괄 제안\n\n유지 확대\n"
    )
    secs = split_markdown_report_sections(md)
    by_id = {s["id"]: s for s in secs}
    assert by_id["summary"]["short"] == "현황"
    assert by_id["urgent"]["short"] == "시급"
    assert by_id["reduce"]["short"] == "축소"
    assert by_id["develop"]["short"] == "개발"
    assert by_id["roadmap"]["short"] == "로드맵"
    assert "본문이 없습니다" not in by_id["urgent"]["markdown"]
    assert _report_section_short_label("제조 시급 품목") == "시급"
    assert _report_section_short_label("공정서 미확보") == "개발"
    assert _report_section_short_label("로드맵 총괄 제안") == "로드맵"


def test_ensure_mandatory_report_sections():
    """요약형 5섹션(현황·시급·축소·개발·로드맵)이 항상 본문을 갖는다."""
    from stock_logic import (
        build_mandatory_section_markdown,
        ensure_mandatory_report_sections,
        split_markdown_report_sections,
        ZERO_STOCK_CATEGORY,
        MANUFACTURE_REDUCE_RECOMMENDATION,
    )

    flags = {
        "dashboard": {
            "kpis": [
                {"key": "managed", "label": "대상품목 수", "value": 3, "display": "3종"}
            ],
            "summary_lines": ["요약"],
        },
        "depletion_category_items": {
            "긴급제조(<1년)": [
                {
                    "name_ko": "소진품A",
                    "manage_no": "D-001",
                    "std_type": "표준생약",
                    "last_qty": 10,
                    "years_left": 0.5,
                    "deplete_ym": "2026년 12월",
                    "depletion_category": "긴급제조(<1년)",
                    "risk_grade": "위험",
                    "priority_score": 80,
                }
            ],
            ZERO_STOCK_CATEGORY: [
                {
                    "name_ko": "제로품",
                    "manage_no": "Z-001",
                    "std_type": "표준생약",
                    "last_qty": 0,
                    "deplete_ym": "-",
                    "depletion_category": ZERO_STOCK_CATEGORY,
                    "risk_grade": "재고없음",
                }
            ],
        },
        "manufacture_candidates": {
            "표준생약": [
                {
                    "name_ko": "검토품",
                    "manage_no": "M-001",
                    "std_type": "표준생약",
                    "last_qty": 50,
                    "years_left": 1.5,
                    "priority_score": 82,
                    "deplete_ym": "2028년 01월",
                    "depletion_category": "제조준비(1~2년)",
                    "risk_grade": "위험",
                }
            ],
            "지표성분": [],
        },
        "manufacture_reduce_items": [
            {
                "name_ko": "과다품",
                "manage_no": "O-001",
                "std_type": "표준생약",
                "coverage_years": 20,
                "recent_rate": 0.2,
                "manufacture_reduce": True,
                "recommendation": MANUFACTURE_REDUCE_RECOMMENDATION,
            }
        ],
        "monitoring_targets": [],
        "missing_compendium_items": [
            {
                "name_ko": "미보유품",
                "origin_ko": "한국",
                "origin_en": "Korea",
                "pharmacopoeia": "KP",
                "pharmacopoeia_kind": "KP",
            }
        ],
        "missing_herb_items": [
            {"name_ko": "미보유품", "origin_count": 1, "pharmacopoeia_kind": "KP"}
        ],
        "compendium_stats": {
            "herb_total": 10,
            "herb_held": 8,
            "herb_missing": 2,
            "herb_coverage_pct": 80.0,
            "origin_total": 12,
            "origin_held": 9,
            "origin_missing": 3,
            "origin_coverage_pct": 75.0,
        },
    }
    match = {
        "stats": flags["compendium_stats"],
        "missing_herb_items": flags["missing_herb_items"],
        "missing_items": flags["missing_compendium_items"],
    }

    bare = "## 기타 분석\n\nAI가 쓴 본문만 있음.\n"
    filled = ensure_mandatory_report_sections(bare, flags=flags, match_result=match)
    assert "종합현황 대시보드" in filled
    assert "소진품A" in filled or "검토품" in filled
    assert "과다품" in filled
    assert "미보유품" in filled
    assert "로드맵 총괄 제안" in filled
    assert "본문이 없습니다" not in filled

    secs = split_markdown_report_sections(filled, flags=flags, match_result=match)
    by_id = {s["id"]: s for s in secs}
    for key in ("summary", "urgent", "reduce", "develop", "roadmap"):
        assert key in by_id, key
        assert "본문이 없습니다" not in by_id[key]["markdown"], key
    assert "소진품A" in by_id["urgent"]["markdown"] or "검토품" in by_id["urgent"]["markdown"]
    assert "과다품" in by_id["reduce"]["markdown"]
    assert "미보유품" in by_id["develop"]["markdown"]
    assert build_mandatory_section_markdown("urgent", flags).startswith("## 제조 시급")


def test_truncated_ai_tables_replaced_with_slim_list():
    """AI 중략 표는 요약형 개발 대상 표로 교체된다."""
    from stock_logic import (
        ensure_mandatory_report_sections,
        split_markdown_report_sections,
    )

    miss_rows = [
        {
            "name_ko": f"미보유{i}",
            "origin_ko": "한국",
            "origin_en": "Korea",
            "pharmacopoeia": "KP",
            "pharmacopoeia_kind": "KP",
        }
        for i in range(1, 6)
    ]
    flags = {
        "dashboard": {
            "kpis": [{"key": "managed", "label": "대상품목 수", "value": 1, "display": "1종"}],
            "summary_lines": ["요약"],
        },
        "depletion_category_items": {},
        "missing_compendium_items": miss_rows,
        "missing_herb_items": [
            {"name_ko": f"미보유{i}", "origin_count": 1, "pharmacopoeia_kind": "KP"}
            for i in range(1, 6)
        ],
        "manufacture_candidates": {},
        "monitoring_targets": [],
        "compendium_stats": {"herb_missing": 5, "origin_missing": 5},
    }
    truncated = (
        "## 공정서 미확보 표준생약 개발 대상\n\n"
        "| # | 생약명(한글) | 기원행수 | 공정서 |\n"
        "| --- | --- | --- | --- |\n"
        "| 1 | 미보유1 | 1 | KP |\n"
        "| ... | (중략) | ... | ... |\n"
        "| 5 | 미보유5 | 1 | KP |\n"
    )
    filled = ensure_mandatory_report_sections(truncated, flags=flags)
    assert "(중략)" not in filled
    assert "미보유3" in filled
    secs = split_markdown_report_sections(filled, flags=flags)
    develop = next(s for s in secs if s["id"] == "develop")
    assert "(중략)" not in develop["markdown"]
    assert "미보유1" in develop["markdown"]


def test_ensure_accel_monitoring_section_always_has_body():
    """요약형 리포트에서도 필수 섹션 본문이 비지 않는다."""
    from stock_logic import (
        ensure_accel_monitoring_in_report,
        format_accel_monitoring_markdown,
        split_markdown_report_sections,
    )

    monitoring = [
        {
            "label": "가속품A (A-001)",
            "name_ko": "가속품A",
            "manage_no": "A-001",
            "std_type": "표준생약",
            "last_qty": 120,
            "acceleration": "급증",
            "acceleration_ratio": 2.5,
            "annual_rate": 40.0,
            "deplete_ym": "2027년 03월",
            "reliability": "A(충분)",
        },
        {
            "label": "증가품B (B-002)",
            "name_ko": "증가품B",
            "manage_no": "B-002",
            "std_type": "지표성분",
            "last_qty": 80,
            "acceleration": "증가",
            "acceleration_ratio": 1.4,
            "annual_rate": 12.0,
            "deplete_ym": "2029년 01월",
            "reliability": "B(보통)",
        },
    ]

    bare = "## 종합현황 대시보드\n\n요약입니다.\n\n## 제조 시급 품목 리스트\n\n소진 내용\n"
    filled = ensure_accel_monitoring_in_report(bare, monitoring)
    assert "본문이 없습니다" not in filled
    md = format_accel_monitoring_markdown(monitoring)
    assert "가속품A" in md and "급증" in md
    assert "| 가속도 |" in md

    secs = split_markdown_report_sections(
        filled,
        flags={
            "monitoring_targets": monitoring,
            "dashboard": {"kpis": [], "summary_lines": []},
        },
    )
    assert {s["id"] for s in secs} >= {"summary", "urgent", "reduce", "develop", "roadmap"}

    filled0 = ensure_accel_monitoring_in_report(bare, [])
    assert "본문이 없습니다" not in filled0
    zero = format_accel_monitoring_markdown([])
    assert "해당 없음" in zero


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
    assert all(r.get("risk_grade") for r in result["표준생약"])
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


def test_deplete_ym_uses_today_not_past_survey_date():
    """소진 예상일시는 마지막 조사일이 아닌 분석 기준일(오늘)부터 산출."""
    item = StockItem(
        manage_no="OLD-01",
        name_ko="과거조사",
        std_type="표준생약",
        corrected_points=[
            StockPoint(date(2015, 1, 1), 300),
            StockPoint(date(2018, 6, 1), 120),
            StockPoint(date(2020, 1, 1), 40),
        ],
    )
    stats = estimate_depletion(item)
    assert stats["years_left"] is not None
    assert stats["deplete_ym"] is not None
    ym = str(stats["deplete_ym"])
    if "장기저수요" in ym or "휴면재고" in ym:
        assert stats["years_left"] >= 15.0 or stats.get("recent_annual_rate") in (None, 0)
        return
    year = int(ym.split("년")[0])
    month = int(ym.split("년")[1].replace("월", "").strip())
    target = date(year, month, 1)
    today = date.today().replace(day=1)
    assert target >= today


def test_acceleration_seven_tiers_and_no_cap():
    from stock_logic import acceleration_from_rates

    new = acceleration_from_rates(0.0, 5.0)
    assert new["label"] == "신규수요"
    assert new["ratio"] is None  # Cap 3.0 제거
    assert new["is_surge"] is True

    surge = acceleration_from_rates(10.0, 25.0)
    assert surge["label"] == "급증" and abs(surge["ratio"] - 2.5) < 1e-9

    up = acceleration_from_rates(10.0, 15.0)
    assert up["label"] == "증가"

    stable = acceleration_from_rates(10.0, 10.0)
    assert stable["label"] == "안정"

    down = acceleration_from_rates(10.0, 6.0)
    assert down["label"] == "감소"

    crash = acceleration_from_rates(10.0, 4.0)
    assert crash["label"] == "급감"

    empty = acceleration_from_rates(0.0, 0.0)
    assert empty["label"] == "분석불가/자료부족"

    low = acceleration_from_rates(0.2, 0.5)
    assert low["label"] == "급증"
    assert low["low_volume_intermittent"] is True


def test_same_year_raw_changes_enable_analysis():
    """동일 연도 내 수량 변동만 있어도 AI 분석 대상·분양속도가 산출되어야 한다.

    연도말 붕괴로 corrected_points가 1개만 남아 has_stock_change=False /
    데이터부족으로 분석이 통째로 스킵되던 회귀를 방지한다.
    """
    from stock_logic import (
        analysis_time_series,
        collect_ai_analysis_flags,
        estimate_depletion,
        items_for_ai_analysis,
        process_excel,
    )

    # 단위 케이스: 2024년 안에서만 감소
    raw = [
        StockPoint(date(2024, 1, 1), 300.0),
        StockPoint(date(2024, 4, 10), 220.0),
        StockPoint(date(2024, 8, 1), 200.0),
    ]
    ye = collapse_to_year_end(raw)
    corr, n = apply_retroactive_correction(raw)
    item = StockItem(
        manage_no="SAME-Y",
        name_ko="동일연도감소",
        std_type="표준생약",
        raw_points=raw,
        year_end_points=ye,
        corrected_points=corr,
        correction_count=n,
    )
    assert len(item.corrected_points) == 1
    assert item.has_stock_change is True
    assert len(analysis_time_series(item)) >= 2
    stats = estimate_depletion(item)
    assert stats["speed"] != "데이터부족"
    assert stats["annual_rate"] is not None and stats["annual_rate"] > 0
    assert stats["years_left"] is not None
    assert items_for_ai_analysis([item]) == [item]

    # 저장소 샘플 엑셀(동일 연도 시계열)도 분석 대상이 0이 아니어야 한다
    sample = Path(__file__).resolve().parents[1] / "sample" / "테스트파일1.xlsx"
    if sample.exists():
        data = process_excel(sample)
        targets = items_for_ai_analysis(data["stock_items"])
        assert len(targets) >= 1
        flags = collect_ai_analysis_flags(data["stock_items"])
        assert len(flags.get("by_code") or {}) >= 1


def test_coverage_long_low_and_manufacture_reduce():
    from stock_logic import (
        LONG_LOW_DEMAND_CATEGORY,
        DORMANT_STOCK_CATEGORY,
        MANUFACTURE_REDUCE_RECOMMENDATION,
        collect_ai_analysis_flags,
        estimate_depletion,
        format_deplete_ym,
        select_long_term_low_items,
    )

    assert format_deplete_ym(20.0, recent_rate=1.0) == LONG_LOW_DEMAND_CATEGORY
    assert format_deplete_ym(None, recent_rate=0.0) == DORMANT_STOCK_CATEGORY

    # 재고 많음 + 최근 저분양 → 장기저수요 + 과다재고 → 제조 축소/보존
    item = StockItem(
        manage_no="LOW-01",
        name_ko="저분양과다",
        std_type="표준생약",
        unit_price=1000,
        corrected_points=[
            StockPoint(date(2018, 1, 1), 500),
            StockPoint(date(2019, 1, 1), 499),
            StockPoint(date(2020, 1, 1), 498),
            StockPoint(date(2021, 1, 1), 497),
            StockPoint(date(2022, 1, 1), 496.5),
            StockPoint(date(2023, 1, 1), 496),
            StockPoint(date(2024, 1, 1), 495.5),
            StockPoint(date(2025, 1, 1), 495),
        ],
    )
    stats = estimate_depletion(item)
    assert stats["long_term_low"] is True
    assert stats["overstock"] is True
    assert stats["manufacture_reduce"] is True
    assert stats["deplete_ym"] == LONG_LOW_DEMAND_CATEGORY
    assert stats["depletion_category"] == LONG_LOW_DEMAND_CATEGORY

    rows = select_long_term_low_items([item])
    assert len(rows) >= 1
    assert rows[0]["recommendation"] == MANUFACTURE_REDUCE_RECOMMENDATION
    flags = collect_ai_analysis_flags([item])
    assert len(flags.get("manufacture_reduce_items") or []) >= 1


def test_manufacture_need_score_excludes_reliability():
    from stock_logic import (
        apply_manufacture_need_scores,
        estimate_depletion,
        get_depletion_stats,
        manufacturing_priority_score,
    )

    # 동일 커버리지·수요·가속도에서 신뢰도만 달라도 점수 동일
    a = manufacturing_priority_score(
        years_left=2.0,
        acceleration_label="급증",
        demand_percentile=0.8,
        pharmacopoeia_weight=1.0,
        reliability_score=1.0,
    )
    b = manufacturing_priority_score(
        years_left=2.0,
        acceleration_label="급증",
        demand_percentile=0.8,
        pharmacopoeia_weight=1.0,
        reliability_score=0.25,
    )
    assert a == b
    assert 0 <= a <= 100

    items = []
    for i, qty in enumerate((80, 60, 40)):
        items.append(
            StockItem(
                manage_no=f"P{i}",
                name_ko=f"점수{i}",
                std_type="표준생약",
                corrected_points=[
                    StockPoint(date(2021, 1, 1), qty + 60),
                    StockPoint(date(2022, 1, 1), qty + 40),
                    StockPoint(date(2023, 1, 1), qty + 20),
                    StockPoint(date(2024, 1, 1), qty),
                ],
            )
        )
    cache = {}
    for it in items:
        get_depletion_stats(it, cache)
    apply_manufacture_need_scores(
        items,
        cache,
        match_result={"by_manage_no": {"P0": "KP", "P1": "KHP"}, "by_label": {}},
    )
    s0 = cache["P0"]["priority_score"]
    s1 = cache["P1"]["priority_score"]
    assert s0 > s1  # KP 가중치
    assert "reliability" in cache["P0"]
    assert cache["P0"]["pharmacopoeia_weight"] == 1.0
    assert cache["P1"]["pharmacopoeia_weight"] == 0.5


def test_compendium_herb_vs_origin_and_hwonchogeun_khp():
    from stock_logic import (
        CompendiumEntry,
        StockItem,
        format_compendium_stats_markdown,
        match_compendium_inventory,
        _normalize_compendium_pharmacopoeia,
    )

    assert _normalize_compendium_pharmacopoeia("훤초근", "") == "KHP"
    assert _normalize_compendium_pharmacopoeia("훤초근", "KP") == "KP"

    entries = [
        CompendiumEntry(name_ko="감초", origin_ko="기원A", pharmacopoeia="KP"),
        CompendiumEntry(name_ko="감초", origin_ko="기원B", pharmacopoeia="KP"),
        CompendiumEntry(name_ko="당귀", origin_ko="기원C", pharmacopoeia="KHP"),
        CompendiumEntry(name_ko="훤초근", origin_ko="기원D", pharmacopoeia=""),
    ]
    items = [StockItem(manage_no="1", name_ko="감초", std_type="표준생약")]
    match = match_compendium_inventory(entries, items)
    st = match["stats"]
    assert st["herb_total"] == 3  # 감초·당귀·훤초근
    assert st["origin_total"] == 4
    assert st["herb_held"] == 1
    assert st["herb_missing"] == 2
    assert st["origin_held"] == 1
    assert st["origin_missing"] == 3
    assert any(e.pharmacopoeia == "KHP" for e in entries if e.name_ko == "훤초근")
    md = format_compendium_stats_markdown(match)
    assert "품목 기준" in md and "기원식물 기준" in md
    assert "품목 확보율" in md and "기원 확보율" in md
    # 미보유를 기원 전체로만 왜곡하지 않음
    assert "미보유 생약 수: **2건**" in md
    assert "미확보 기원 수: **3건**" in md


def test_strip_duplicate_auto_summary_opinion():
    from stock_logic import _strip_auto_summary_opinion_blocks, ensure_mandatory_report_sections

    raw = (
        "## 1페이지 요약 대시보드\n\n### 자동 종합 의견\n- 중복 A\n\n"
        "## 기타\n\n### 자동 종합 의견\n- 중복 B\n"
    )
    stripped = _strip_auto_summary_opinion_blocks(raw)
    assert "자동 종합 의견" not in stripped
    assert "중복 A" not in stripped
    assert "## 기타" in stripped

    flags = {
        "dashboard": {
            "kpis": [{"label": "대상품목 수", "display": "1종"}],
            "summary_lines": ["단일 요약"],
        },
        "depletion_category_items": {},
        "manufacture_candidates": {"표준생약": [], "지표성분": []},
        "monitoring_targets": [],
    }
    filled = ensure_mandatory_report_sections(raw, flags=flags)
    assert filled.count("### 자동 종합 의견") == 1
    assert "단일 요약" in filled


def test_manufacture_section_always_has_priority_formula():
    from stock_logic import PRIORITY_FORMULA_KO, format_manufacture_review_markdown

    md = format_manufacture_review_markdown({"표준생약": [], "지표성분": []})
    assert "제조 우선순위" in md
    assert PRIORITY_FORMULA_KO.splitlines()[0] in md


def test_unique_missing_compendium_examples():
    from stock_logic import _unique_missing_compendium_examples

    rows = [
        {"name_ko": "감초", "pharmacopoeia_kind": "KP"},
        {"name_ko": "감초", "pharmacopoeia_kind": "KHP"},
        {"name_ko": "당귀"},
    ]
    assert _unique_missing_compendium_examples(rows) == ["감초(KP)", "당귀"]


def test_depletion_stats_cache_deduplicates():
    """collect_ai_analysis_flags가 품목당 estimate_depletion을 1회만 호출."""
    import stock_logic
    from unittest.mock import patch

    path = _make_sample_xlsx()
    try:
        _, items = stock_logic.load_stock_excel(path)
        calls: list[str] = []
        orig = stock_logic.estimate_depletion

        def spy(it):
            calls.append(stock_logic._stock_item_stats_key(it))
            return orig(it)

        with patch.object(stock_logic, "estimate_depletion", spy):
            flags = stock_logic.collect_ai_analysis_flags(items)

        assert len(calls) == len(set(calls))
        assert flags["by_code"]
        assert flags["dashboard"]["kpis"]
    finally:
        os.remove(path)


def test_excel_engine_branch_and_badzip_fallback():
    import zipfile
    from pathlib import Path
    from unittest.mock import patch

    from stock_logic import _excel_engine_for_path, read_excel_dataframe

    assert _excel_engine_for_path("a.xls") == "xlrd"
    assert _excel_engine_for_path("a.XLS") == "xlrd"
    assert _excel_engine_for_path("a.xlsx") == "openpyxl"

    path = _make_sample_xlsx()
    try:
        df = read_excel_dataframe(path)
        assert "관리번호" in df.columns

        # openpyxl이 BadZipFile을 내면 xlrd로 재시도
        calls: list[str] = []

        def fake_read(p, engine=None, **kwargs):
            calls.append(engine or "")
            if engine == "openpyxl":
                raise zipfile.BadZipFile("File is not a zip file")
            if engine == "xlrd":
                return pd.DataFrame([{"관리번호": "X", "한글명": "테스트"}])
            raise AssertionError(f"unexpected engine {engine}")

        with patch("stock_logic.pd.read_excel", side_effect=fake_read):
            out = read_excel_dataframe("/tmp/fake.xlsx")
        assert calls == ["openpyxl", "xlrd"]
        assert list(out.columns) == ["관리번호", "한글명"]
    finally:
        os.remove(path)


def test_process_excels_skips_failed_file():
    from pathlib import Path

    from stock_logic import process_excels

    good = _make_sample_xlsx()
    bad = tempfile.mkstemp(suffix=".xlsx")[1]
    try:
        with open(bad, "wb") as f:
            f.write(b"not-an-excel")
        result = process_excels([good, bad])
        assert result["row_count"] >= 1
        assert len(result["file_paths"]) == 1
        assert Path(result["file_paths"][0]).name == Path(good).name
        assert len(result["failed_files"]) == 1
        assert result["failed_files"][0]["name"] == Path(bad).name
    finally:
        os.remove(good)
        os.remove(bad)


def test_normalize_excel_path_filters_temp_and_empty():
    from stock_logic import normalize_excel_path

    assert normalize_excel_path(None) is None
    assert normalize_excel_path(("/tmp/a.xlsx",)) is None  # missing file
    fd, empty = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    try:
        assert os.path.getsize(empty) == 0
        assert normalize_excel_path(empty) is None
        assert normalize_excel_path(os.path.join(os.path.dirname(empty), "~$book.xlsx")) is None
    finally:
        os.remove(empty)
    good = _make_sample_xlsx()
    try:
        assert normalize_excel_path(good) is not None
        assert isinstance(normalize_excel_path(good), str)
    finally:
        os.remove(good)



def test_name_ko_stock_grouping_and_zero_filter():
    """동일 한글명 로트 중 하나라도 재고가 있으면 품목은 재고 보유."""
    from stock_logic import (
        StockItem,
        StockPoint,
        build_name_ko_stock_map,
        group_by_depletion_category_items,
        is_name_level_zero_stock,
        name_ko_group_has_stock,
        ZERO_STOCK_CATEGORY,
    )
    from datetime import date

    def _item(code, name, qty, *, start=10.0):
        it = StockItem(manage_no=code, name_ko=name, std_type="표준생약", balance=qty)
        # AI 대상은 수량 변동 이력이 있어야 함
        it.corrected_points = [
            StockPoint(change_date=date(2022, 12, 31), quantity=float(start)),
            StockPoint(change_date=date(2024, 12, 31), quantity=float(qty)),
        ]
        return it

    items = [
        _item("B-1", "베타인", 0, start=5),
        _item("B-2", "베타인", 3, start=8),
        _item("Z-1", "제로만", 0, start=4),
    ]
    m = build_name_ko_stock_map(items)
    assert m["베타인"]["has_stock"] is True
    assert m["베타인"]["total_qty"] == 3
    assert m["제로만"]["has_stock"] is False
    assert name_ko_group_has_stock(items[0], m) is True
    assert is_name_level_zero_stock(items[0], m) is False
    assert is_name_level_zero_stock(items[2], m) is True

    cats = group_by_depletion_category_items(items, name_ko_map=m)
    zero_names = {r.get("name_ko") for r in cats.get(ZERO_STOCK_CATEGORY) or []}
    assert "베타인" not in zero_names
    assert "제로만" in zero_names


def test_name_en_inventory_group_compendium_match():
    """영문명 동일 로트는 공정서 매칭·통계에서 1건으로 집계."""
    from stock_logic import (
        CompendiumEntry,
        StockItem,
        attach_compendium_match_to_flags,
        build_name_en_inventory_groups,
        collect_ai_analysis_flags,
        match_compendium_inventory,
    )

    entries = [
        CompendiumEntry(name_ko="노포", name_en="Nofo Extract", pharmacopoeia="KP"),
        CompendiumEntry(name_ko="감초", name_en="Glycyrrhizae Radix", pharmacopoeia="KP"),
    ]
    items = [
        StockItem(manage_no="NOFO2009", name_ko="노포A", name_en="Nofo Extract", balance=1),
        StockItem(manage_no="NOFO2023", name_ko="노포B", name_en="Nofo Extract", balance=2),
        StockItem(manage_no="G1", name_ko="감초", name_en="Glycyrrhizae Radix", balance=5),
    ]
    groups = build_name_en_inventory_groups(items)
    assert len(groups) == 2
    match = match_compendium_inventory(entries, items)
    assert match["stats"]["unique_inventory_groups"] == 2
    assert match["stats"]["inventory_matched"] == 2
    assert match["stats"]["inventory_matched_lots"] == 3
    assert len(match.get("name_en_match_map") or {}) == 2

    flags = collect_ai_analysis_flags(items)
    flags = attach_compendium_match_to_flags(flags, match)
    maps = flags.get("chat_analysis_maps") or {}
    assert maps.get("name_en_match_map")
    assert maps.get("name_ko_stock_map")


def test_report_nav_has_no_other_tab():
    from stock_logic import split_markdown_report_sections

    md = (
        "## 종합현황 대시보드\n\n요약본문\n\n"
        "## 제조 시급 품목 리스트\n\n소진본문\n\n"
        "## 예기치 않은 잔여 섹션\n\n잔여본문\n"
    )
    secs = split_markdown_report_sections(md)
    assert all(s.get("id") != "other" for s in secs)
    assert all(s.get("short") != "기타" for s in secs)
    summary = next(s for s in secs if s["id"] == "summary")
    assert "잔여본문" in summary["markdown"]


def test_chatbot_prompt_persona_and_maps_json():
    from stock_logic import (
        StockItem,
        build_ai_prompt,
        build_followup_prompt,
        collect_ai_analysis_flags,
        format_chat_analysis_maps_json,
    )

    items = [StockItem(manage_no="1", name_ko="감초", balance=10)]
    flags = collect_ai_analysis_flags(items)
    j = format_chat_analysis_maps_json(flags)
    assert "구조화 분석 맵 JSON" in j
    assert "name_ko_stock_summary" in j
    prompt = build_ai_prompt(items, flags=flags)
    assert "수석 데이터 분석가" in prompt or "생약표준품" in prompt
    assert "현황 수치" in prompt and ("권고" in prompt or "원인" in prompt)
    assert "Token Diet" in prompt
    assert "구조화 분석 맵 JSON" not in prompt
    fu = build_followup_prompt("초기", "위험 품목은?", items=items, flags=flags)
    assert "최고 수석 데이터 분석가" in fu or "수석 데이터 분석가" in fu
    assert "권고 액션 플랜" in fu or "권고" in fu
    assert "구조화 분석 맵 JSON" in fu


def test_export_markdown_report_to_docx(tmp_path=None):
    import tempfile
    from pathlib import Path
    from docx import Document
    from stock_logic import export_markdown_report_to_docx

    md = (
        "## 요약\n\n**중요** 본문입니다.\n\n"
        "- 항목 하나\n"
        "- 항목 둘\n\n"
        "| 이름 | 값 |\n| --- | --- |\n| 감초 | 1 |\n"
    )
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "report.docx"
        export_markdown_report_to_docx(md, out)
        assert out.exists() and out.stat().st_size > 500
        doc = Document(str(out))
        assert len(doc.tables) >= 1
        tbl = doc.tables[0]
        assert tbl.cell(0, 0).text.strip() == "이름"
        assert "감초" in tbl.cell(1, 0).text
        assert doc.styles["Normal"].font.name == "맑은 고딕"

    with tempfile.TemporaryDirectory() as td2:
        try:
            export_markdown_report_to_docx("   ", Path(td2) / "empty.docx")
            assert False, "expected ValueError"
        except ValueError:
            pass


def test_export_docx_markdown_table_as_grid():
    """마크다운 표가 Word Table Grid로 생성되고 구분선은 제외."""
    import tempfile
    from pathlib import Path
    from docx import Document
    from stock_logic import export_markdown_report_to_docx

    md = (
        "### 소진\n\n"
        "| 한글명 | 관리번호 | 재고 |\n"
        "| --- | --- | --- |\n"
        "| 감초 | STD-001 | 10 |\n"
        "| 황기 | STD-002 | 5 |\n\n"
        "본문 이어짐.\n"
    )
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "tbl.docx"
        export_markdown_report_to_docx(md, out)
        doc = Document(str(out))
        assert len(doc.tables) == 1
        t = doc.tables[0]
        assert len(t.rows) == 3  # header + 2 data (separator skipped)
        assert t.cell(0, 0).text.strip() == "한글명"
        assert t.cell(1, 0).text.strip() == "감초"
        assert t.cell(2, 1).text.strip() == "STD-002"
        # 헤더 bold
        assert any(r.bold for r in t.cell(0, 0).paragraphs[0].runs)


def test_export_docx_handles_large_table_as_plain_lines():
    """대량 표는 상한으로 잘라 Table Grid로 저장(프리징 방지)."""
    import tempfile
    from pathlib import Path
    from docx import Document
    from stock_logic import export_markdown_report_to_docx

    rows = ["| a | b |", "| --- | --- |"] + [f"| r{i} | v{i} |" for i in range(500)]
    md = "## 표\n\n" + "\n".join(rows)
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "big.docx"
        export_markdown_report_to_docx(md, out)
        doc = Document(str(out))
        assert len(doc.tables) == 1
        # MAX_TABLE_ROWS=80 + 상한 안내 행
        assert len(doc.tables[0].rows) == 81
        assert "상한" in doc.tables[0].rows[-1].cells[0].text


def test_slim_report_five_sections_no_full_catalog_blowup():
    """요약형 리포트는 5섹션이며 전수 장문 헤딩을 주입하지 않는다."""
    from stock_logic import (
        ensure_mandatory_report_sections,
        format_appendix_full_catalog_markdown,
        INCLUDE_REPORT_APPENDIX_DEFAULT,
    )

    flags = {
        "dashboard": {
            "kpis": [{"key": "managed", "label": "대상품목 수", "value": 2, "display": "2종"}],
            "summary_lines": ["의견"],
        },
        "depletion_category_items": {
            "긴급제조(<1년)": [
                {"name_ko": f"시급{i}", "manage_no": f"U{i}", "years_left": 0.5, "priority_score": 90 - i}
                for i in range(30)
            ]
        },
        "manufacture_candidates": {},
        "monitoring_targets": [{"name_ko": "가속1", "acceleration": "급증"}],
        "missing_herb_items": [
            {"name_ko": f"미보{i}", "pharmacopoeia_kind": "KP", "origin_count": 1}
            for i in range(40)
        ],
        "compendium_stats": {"herb_missing": 40, "origin_missing": 40},
    }
    filled = ensure_mandatory_report_sections("", flags=flags)
    assert "## 종합현황 대시보드" in filled
    assert "## 제조 시급 품목 리스트" in filled
    assert "## 로드맵 총괄 제안" in filled
    assert INCLUDE_REPORT_APPENDIX_DEFAULT is False
    assert "부록(별첨) — 전수 목록" not in filled or "생략" in filled
    # TOP20만
    assert "시급19" in filled
    assert "시급29" not in filled
    appendix = format_appendix_full_catalog_markdown(flags)
    assert "부록" in appendix


def test_thin_roadmap_replaced_with_pillars_and_examples():
    """AI가 로드맵을 한 줄만 쓰면 유지/확대/신규/축소 + 예시 품명으로 교체."""
    from stock_logic import (
        _roadmap_section_is_thin,
        ensure_mandatory_report_sections,
        format_roadmap_markdown,
        split_markdown_report_sections,
    )

    flags = {
        "dashboard": {
            "kpis": [{"key": "managed", "label": "대상", "value": 3, "display": "3종"}],
            "summary_lines": ["요약"],
        },
        "by_code": {"A": {}},
        "depletion_category_items": {
            "긴급제조(<1년)": [
                {
                    "manage_no": "A",
                    "name_ko": "감초",
                    "years_left": 0.5,
                    "priority_score": 90,
                }
            ],
            "적정(5~10년)": [{"manage_no": "C", "name_ko": "작약"}],
            "모니터링(3~5년)": [{"manage_no": "D", "name_ko": "천궁"}],
        },
        "manufacture_reduce_items": [
            {
                "manage_no": "B",
                "name_ko": "당귀",
                "manufacture_reduce": True,
                "stock_value": 1,
            }
        ],
        "long_term_low_items": [],
        "monitoring_targets": [{"name_ko": "인삼", "acceleration": "급증"}],
        "missing_herb_items": [{"name_ko": "황기", "pharmacopoeia_kind": "KP"}],
    }
    thin = (
        "## 로드맵 총괄 제안\n\n"
        "유지 · 확대 · 신규 · 축소 관점의 실행 로드맵 (요약형).\n"
    )
    assert _roadmap_section_is_thin(thin) is True
    filled = ensure_mandatory_report_sections(thin, flags=flags)
    for pillar in ("### 유지", "### 확대", "### 신규", "### 축소"):
        assert pillar in filled
    assert "감초" in filled and "당귀" in filled
    secs = split_markdown_report_sections(filled, flags=flags)
    rm = next(s for s in secs if s["id"] == "roadmap")
    assert "### 유지" in rm["markdown"]
    assert "감초" in rm["markdown"]
    full = format_roadmap_markdown(flags)
    assert _roadmap_section_is_thin(full) is False



def test_compendium_en_first_match_and_failures():
    """영문명 우선 매칭 + 미매칭은 match_failures로 분리."""
    from stock_logic import CompendiumEntry, StockItem, match_compendium_inventory

    entries = [
        CompendiumEntry(name_ko="감초", name_en="Glycyrrhizae Radix", pharmacopoeia="KP"),
        CompendiumEntry(name_ko="당귀", name_en="Angelicae Gigantis Radix", pharmacopoeia="KP"),
    ]
    items = [
        # 한글명은 다르지만 영문명으로 매칭되어야 함
        StockItem(manage_no="G1", name_ko="감초추출물", name_en="Glycyrrhizae Radix", balance=5),
        StockItem(manage_no="U1", name_ko="미등록생약", name_en="Unknown Herba", balance=2),
    ]
    match = match_compendium_inventory(entries, items)
    types = [c["match_type"] for c in match["corrections"]]
    assert "fuzzy_en" in types
    assert match["stats"]["match_failure_count"] == 1
    assert len(match["match_failures"]) == 1
    assert match["match_failures"][0]["name_en"] == "Unknown Herba"
    assert match["stats"]["missing_count"] == 1  # 당귀 미보유


def test_compendium_origin_exception_excluded_from_missing():
    from stock_logic import (
        CompendiumEntry,
        StockItem,
        format_compendium_stats_markdown,
        match_compendium_inventory,
    )

    entries = [
        CompendiumEntry(name_ko="감초", name_en="Glycyrrhizae Radix", pharmacopoeia="KP"),
        CompendiumEntry(
            name_ko="예외약",
            name_en="Excepta Herba",
            pharmacopoeia="KP",
            origin_ko="그 변종",
            is_origin_exception=True,
        ),
    ]
    items = [StockItem(manage_no="G1", name_ko="감초", name_en="Glycyrrhizae Radix", balance=1)]
    match = match_compendium_inventory(entries, items)
    assert match["stats"]["origin_exception_count"] == 1
    assert all(e.name_ko != "예외약" for e in match["missing"])
    assert match["stats"]["missing_count"] == 0
    md = format_compendium_stats_markdown(match)
    assert "기원 예외" in md


def test_compendium_std_ref_ab_groups_from_columns():
    import pandas as pd
    from stock_logic import _parse_compendium_entries, format_compendium_stats_markdown, match_compendium_inventory, StockItem

    df = pd.DataFrame(
        [
            {
                "생약명(한글)": "감초",
                "생약명(영어)": "Glycyrrhizae Radix",
                "기원(한글)": "콩과",
                "확인시험 표준품 1": "감초표준품",
                "정량법 표준품1": "",
                "공정서": "KP",
            },
            {
                "생약명(한글)": "황기",
                "생약명(영어)": "Astragali Radix",
                "기원(한글)": "콩과",
                "확인시험 표준품 1": "",
                "정량법 표준품1": "",
                "공정서": "KP",
            },
        ]
    )
    entries = _parse_compendium_entries(df)
    assert entries[0].std_group == "A" and entries[0].std_ox == "O"
    assert entries[1].std_group == "B" and entries[1].std_ox == "X"
    items = [StockItem(manage_no="1", name_ko="감초", name_en="Glycyrrhizae Radix", balance=1)]
    match = match_compendium_inventory(entries, items)
    assert match["stats"]["std_group_a_count"] == 1
    assert match["stats"]["std_group_b_count"] == 1
    md = format_compendium_stats_markdown(match)
    assert "표준품 명시" in md and "(A)" in md


def test_successor_lot_excluded_from_manufacture_and_urgent():
    from datetime import date
    from stock_logic import (
        StockItem,
        StockPoint,
        build_lot_succession_map,
        group_by_depletion_category_items,
        is_superseded_lot,
        select_manufacture_candidates,
        select_monitoring_targets,
    )

    old = StockItem(
        manage_no="LOT2013",
        name_ko="작약",
        name_en="Paeoniae Radix",
        balance=5,
        std_type="표준생약",
        registered_date=date(2013, 1, 1),
    )
    new = StockItem(
        manage_no="LOT2015",
        name_ko="작약",
        name_en="Paeoniae Radix",
        balance=40,
        std_type="표준생약",
        registered_date=date(2015, 1, 1),
    )
    old.corrected_points = [
        StockPoint(date(2013, 1, 1), 200),
        StockPoint(date(2014, 1, 1), 80),
        StockPoint(date(2015, 1, 1), 5),
    ]
    new.corrected_points = [
        StockPoint(date(2015, 1, 1), 100),
        StockPoint(date(2016, 1, 1), 70),
        StockPoint(date(2017, 1, 1), 40),
    ]
    succ = build_lot_succession_map([old, new])
    assert is_superseded_lot(old, succ)
    assert not is_superseded_lot(new, succ)

    mfg = select_manufacture_candidates([old, new])
    mfg_nos = {r["manage_no"] for rows in mfg.values() for r in rows}
    assert "LOT2013" not in mfg_nos
    assert "LOT2015" in mfg_nos

    mon = select_monitoring_targets([old, new])
    mon_nos = {r["manage_no"] for r in mon}
    assert "LOT2013" not in mon_nos

    cats = group_by_depletion_category_items([old, new])
    urgent = cats.get("긴급제조(<1년)") or []
    urgent_nos = {r.get("manage_no") for r in urgent}
    assert "LOT2013" not in urgent_nos


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
        test_ensure_mandatory_report_sections,
        test_truncated_ai_tables_replaced_with_slim_list,
        test_ensure_accel_monitoring_section_always_has_body,
        test_manufacture_candidates_always_top10_by_score,
        test_compendium_missing_set_and_followup_filter,
        test_compendium_db_not_timeseries,
        test_year_display_and_continuous_axis,
        test_multi_file_merge,
        test_scatter_cat_label,
        test_scatter3d_records_from_sample,
        test_deplete_ym_uses_today_not_past_survey_date,
        test_acceleration_seven_tiers_and_no_cap,
        test_same_year_raw_changes_enable_analysis,
        test_coverage_long_low_and_manufacture_reduce,
        test_manufacture_need_score_excludes_reliability,
        test_compendium_herb_vs_origin_and_hwonchogeun_khp,
        test_strip_duplicate_auto_summary_opinion,
        test_manufacture_section_always_has_priority_formula,
        test_unique_missing_compendium_examples,
        test_depletion_stats_cache_deduplicates,
        test_excel_engine_branch_and_badzip_fallback,
        test_process_excels_skips_failed_file,
        test_normalize_excel_path_filters_temp_and_empty,
        test_name_ko_stock_grouping_and_zero_filter,
        test_name_en_inventory_group_compendium_match,
        test_report_nav_has_no_other_tab,
        test_chatbot_prompt_persona_and_maps_json,
        test_export_markdown_report_to_docx,
        test_export_docx_markdown_table_as_grid,
        test_export_docx_handles_large_table_as_plain_lines,
        test_slim_report_five_sections_no_full_catalog_blowup,
        test_thin_roadmap_replaced_with_pillars_and_examples,
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
