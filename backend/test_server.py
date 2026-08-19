import tempfile
import unittest
from pathlib import Path

from backend.server import CollaborationStore, validate_value
from backend.workbook_v2 import SparseWorkbookStore


class ValueValidationTests(unittest.TestCase):
    def test_every_column_accepts_arbitrary_text(self) -> None:
        self.assertEqual(validate_value("publish_count", "=SUM(A1:A3)"), "=SUM(A1:A3)")
        self.assertEqual(validate_value("openings", -1), "-1")
        self.assertEqual(validate_value("teacher", "  王老师  "), "  王老师  ")
        self.assertEqual(validate_value("account_name", ""), "")


class StoreTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.store = CollaborationStore(Path(self.temporary.name) / "state.json")

    def tearDown(self) -> None:
        self.temporary.cleanup()

    def test_cell_updates_are_versioned_and_attributed(self) -> None:
        result = self.store.update_cell(2, "publish_count", 8, 0, "测试员", "client-12345678")
        self.assertFalse(result["conflict"])
        self.assertEqual(result["meta"]["version"], 1)
        self.assertEqual(result["meta"]["updated_by"], "测试员")
        conflict = self.store.update_cell(2, "publish_count", 9, 0, "另一人", "client-87654321")
        self.assertTrue(conflict["conflict"])
        self.assertEqual(conflict["value"], "8")

    def test_new_row_is_blank_and_has_no_business_label(self) -> None:
        result = self.store.add_row("测试员", "client-12345678")
        self.assertIn("account_name", result["row"])
        self.assertEqual(result["row"]["account_name"], "")
        self.assertIn("short_graphic_leads", result["row"])
        self.assertNotIn("formula", result["row"])
        self.assertTrue(all(value in ("", None) for key, value in result["row"].items() if key != "id"))
        self.assertEqual(result["activity"]["field_label"], "新增行")

    def test_batch_paste_is_atomic_and_versioned(self) -> None:
        events = self.store.batch_update_cells(
            [
                {"row_id": 2, "field": "publish_count", "value": "11", "version": 0},
                {"row_id": 2, "field": "total_followers", "value": "9001", "version": 0},
            ],
            "粘贴测试",
            "client-12345678",
        )
        self.assertEqual(len(events), 2)
        snapshot = self.store.snapshot()
        row = next(item for item in snapshot["rows"] if item["id"] == 2)
        self.assertEqual(row["publish_count"], "11")
        self.assertEqual(row["total_followers"], "9001")

    def test_batch_preserves_formula_as_plain_text(self) -> None:
        events = self.store.batch_update_cells(
            [{"row_id": 2, "field": "publish_count", "value": "=1+1", "version": 0}],
            "粘贴测试",
            "client-12345678",
        )
        self.assertEqual(events[0]["value"], "=1+1")


class SparseWorkbookTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.store = SparseWorkbookStore(Path(self.temporary.name) / "workbook-v2.json")

    def tearDown(self) -> None:
        self.temporary.cleanup()

    def test_far_coordinate_is_sparse_and_expands_used_range(self) -> None:
        result = self.store.apply_operations(
            "sheet-1",
            [{"type": "cells", "cells": [{"row": 100_000, "col": 10_000, "value": "远端内容"}]}],
            "测试员",
            "client-12345678",
        )
        self.assertEqual(result["used_range"]["bottom"], 100_000)
        self.assertEqual(result["used_range"]["right"], 10_000)
        self.assertEqual(result["content_range"]["bottom"], 100_000)
        self.assertEqual(result["content_range"]["right"], 10_000)
        snapshot = self.store.snapshot("sheet-1", 100_000, 10_000, 1, 1)
        self.assertEqual(snapshot["sheet"]["cells"]["100000:10000"], "远端内容")

    def test_window_read_does_not_return_far_sparse_cells(self) -> None:
        self.store.apply_operations(
            "sheet-1",
            [{"type": "cells", "cells": [{"row": 0, "col": 0, "value": "首屏"}, {"row": 100_000, "col": 10_000, "value": "远端"}]}],
            "测试员",
            "client-12345678",
        )
        snapshot = self.store.snapshot("sheet-1", 0, 0, 24, 12)
        self.assertEqual(snapshot["sheet"]["content_range"], {"top": 0, "left": 0, "bottom": 100_000, "right": 10_000})
        self.assertEqual(snapshot["sheet"]["cells"], {"0:0": "首屏"})

    def test_window_retains_sparse_global_axis_metrics(self) -> None:
        self.store.apply_operations(
            "sheet-1",
            [
                {"type": "layout", "action": "row_height", "payload": {"row": 80_000, "height": 58}},
                {"type": "layout", "action": "column_width", "payload": {"col": 70_000, "width": 180}},
            ],
            "测试员",
            "client-12345678",
        )
        snapshot = self.store.snapshot("sheet-1", 0, 0, 24, 12)
        self.assertEqual(snapshot["sheet"]["row_heights"], {"80000": 58})
        self.assertEqual(snapshot["sheet"]["column_widths"], {"70000": 180})

    def test_clearing_last_value_contracts_used_range(self) -> None:
        self.store.apply_operations(
            "sheet-1",
            [{"type": "cells", "cells": [{"row": 80, "col": 40, "value": "x"}]}],
            "测试员",
            "client-12345678",
        )
        result = self.store.apply_operations(
            "sheet-1",
            [{"type": "cells", "cells": [{"row": 80, "col": 40, "value": ""}]}],
            "测试员",
            "client-12345678",
        )
        self.assertEqual(result["used_range"], {"top": 0, "left": 0, "bottom": -1, "right": -1})
        self.assertEqual(result["content_range"], {"top": 0, "left": 0, "bottom": -1, "right": -1})

    def test_structure_counts_as_used_range_and_columns_are_not_fixed(self) -> None:
        result = self.store.apply_operations(
            "sheet-1",
            [{"type": "layout", "action": "column_width", "payload": {"col": 500_000, "width": 160}}],
            "测试员",
            "client-12345678",
        )
        self.assertEqual(result["used_range"]["right"], 500_000)
        self.assertEqual(result["used_range"]["bottom"], 0)
        self.assertEqual(result["content_range"], {"top": 0, "left": 0, "bottom": -1, "right": -1})

    def test_formatting_does_not_expand_scroll_content_range(self) -> None:
        result = self.store.apply_operations(
            "sheet-1",
            [{"type": "format", "cells": [{"row": 3_039, "col": 99}], "format": {"align": "left"}}],
            "测试员",
            "client-12345678",
        )
        self.assertEqual(result["used_range"], {"top": 0, "left": 0, "bottom": 3_039, "right": 99})
        self.assertEqual(result["content_range"], {"top": 0, "left": 0, "bottom": -1, "right": -1})

    def test_legacy_migration_recomputes_content_range_without_layout_only_extent(self) -> None:
        legacy_file = Path(self.temporary.name) / "state.json"
        legacy_file.write_text(
            """{
              "schema_version": 1,
              "rows": [{"id": 7, "account_name": "内容", "publish_count": "=SUM(A1:A1)"}],
              "layout": {
                "formats": {"7:custom_1": {"fill": "#ddebf7"}},
                "row_heights": {},
                "column_widths": {},
                "merges": []
              }
            }""",
            encoding="utf-8",
        )
        migrated = SparseWorkbookStore(Path(self.temporary.name) / "migrated-v2.json", legacy_file)
        snapshot = migrated.snapshot("sheet-1", 0, 0, 16, 16)
        self.assertEqual(snapshot["sheet"]["content_range"], {"top": 0, "left": 0, "bottom": 0, "right": 3})
        self.assertEqual(snapshot["sheet"]["used_range"], {"top": 0, "left": 0, "bottom": 0, "right": 9})
        self.assertEqual(snapshot["sheet"]["cells"]["0:3"], "=SUM(A1:A1)")


if __name__ == "__main__":
    unittest.main()
