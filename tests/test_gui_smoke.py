"""헤드리스 GUI 스모크 테스트 (QT_QPA_PLATFORM=offscreen 권장)."""

from __future__ import annotations

import os
import sys
import tempfile
from datetime import date

import pandas as pd
from PyQt6.QtWidgets import QApplication

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from main import MainWindow  # noqa: E402
from stock_logic import process_excel  # noqa: E402


def _sample_path() -> str:
    df = pd.DataFrame(
        [
            {
                "순번": 1,
                "표준품구분": "생약(표준생약)",
                "관리번호": "STD-001",
                "한글명": "감초",
                "영문명": "G",
                "잔고": 120,
                "등록일자": "2020-01-01",
                "분양여부": "Y",
                "규격": "1g",
                "변경일자1": date(2022, 1, 1),
                "재고량1": 150,
                "변경일자2": date(2024, 1, 1),
                "재고량2": 120,
            }
        ]
    )
    fd, path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    df.to_excel(path, index=False)
    return path


def main() -> int:
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    app = QApplication.instance() or QApplication(sys.argv)
    win = MainWindow()

    path = _sample_path()
    try:
        data = process_excel(path)
        win._on_excel_loaded(data)

        assert win.table.rowCount() == 1
        assert "재고" in [
            win.table.horizontalHeaderItem(i).text() for i in range(win.table.columnCount())
        ]
        assert "등록일자" not in [
            win.table.horizontalHeaderItem(i).text() for i in range(win.table.columnCount())
        ]
        assert win.category_combo.findText("표준생약") >= 0
        assert win.item_combo.count() >= 2
        assert win.statusBar().findChild(type(win.status_excel)).text()  # status exists
        assert any(
            getattr(w, "text", lambda: "")() == "made by 2026MFDSyouthinternKYHLCY"
            for w in win.statusBar().children()
            if hasattr(w, "text")
        )
        assert win.tabs.count() == 4
        print("PASS GUI smoke")
        win.close()
        return 0
    finally:
        os.remove(path)


if __name__ == "__main__":
    raise SystemExit(main())
