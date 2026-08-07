"""헤드리스 GUI 스모크 테스트 (QT_QPA_PLATFORM=offscreen 권장)."""

from __future__ import annotations

import os
import sys

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from PySide6.QtWidgets import QApplication

from main import StockApp


SAMPLE = os.path.join(ROOT, "sample", "테스트파일1.xlsx")


def main() -> int:
    app = QApplication.instance() or QApplication(sys.argv)
    win = StockApp()
    win.load_file(SAMPLE)

    assert win.table.rowCount() == 5
    assert win.item_combo.count() == 5
    assert "STD-001" in win.item_combo.itemText(0)

    # 드롭존 / 상태바
    assert "테스트파일1.xlsx" in win.drop_zone.file_label.text()
    assert "테스트파일1.xlsx" in win.status_file.text()
    assert "소급 보정" in win.status_correction.text()
    assert win._correction_count >= 1  # STD-002 동일일자 보정

    # API 키 상태
    win.api_key_edit.setText("")
    win.refresh_api_status()
    assert "미연결" in win.status_api.text()
    win.api_key_edit.setText("dummy-key")
    assert "연결 준비됨" in win.status_api.text()

    # 표 더블클릭 → 차트 탭 + 해당 관리번호 선택
    win.on_table_double_clicked(1, 0)
    assert win.tabs.currentIndex() == 1
    assert win.item_combo.currentData() == "NST-014"
    assert "NST-014" in win.chart_info.text()

    # 콤보로 STD-003 선택
    idx = next(i for i in range(win.item_combo.count()) if win.item_combo.itemData(i) == "STD-003")
    win.item_combo.setCurrentIndex(idx)
    assert "STD-003" in win.chart_info.text()
    assert "300" in win.chart_info.text() and "200" in win.chart_info.text()

    # AI 프롬프트에 변화 품목만
    prompt = win.ai_prompt_view.toPlainText()
    assert "STD-001" in prompt
    assert "STD-003" in prompt
    assert "STD-002" not in prompt
    assert "NST-020" not in prompt
    assert "3건" in win.ai_summary.text()

    print("PASS GUI smoke")
    win.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
