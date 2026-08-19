"""Sparse, coordinate-based workbook storage for the infinite GridBook grid.

The original demo persists a dense list of business rows and a fixed set of
columns.  This module deliberately uses a separate V2 JSON file and numeric
coordinates so browsing empty space never allocates or persists blank cells.
"""

from __future__ import annotations

import json
import os
import tempfile
import threading
import uuid
from copy import deepcopy
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable


DEFAULT_ROW_HEIGHT = 29
DEFAULT_COLUMN_WIDTH = 120
MAX_COORDINATE = 2_000_000_000
MAX_BATCH_CELLS = 100_000


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")


def coordinate_key(row: int, col: int) -> str:
    return f"{row}:{col}"


def parse_coordinate_key(key: str) -> tuple[int, int]:
    row, col = key.split(":", 1)
    return int(row), int(col)


def _coordinate(value: Any, label: str) -> int:
    try:
        normalized = int(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"无效的{label}") from exc
    if normalized < 0 or normalized > MAX_COORDINATE:
        raise ValueError(f"{label}超出允许范围")
    return normalized


def _positive(value: Any, label: str, maximum: int = MAX_BATCH_CELLS) -> int:
    try:
        normalized = int(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"无效的{label}") from exc
    if normalized < 1 or normalized > maximum:
        raise ValueError(f"{label}超出允许范围")
    return normalized


def _empty_sheet(identifier: str, name: str) -> dict[str, Any]:
    return {
        "id": identifier,
        "name": name,
        "cells": {},
        "meta": {},
        "formats": {},
        "row_heights": {},
        "column_widths": {},
        "merges": [],
        "used_range": {"top": 0, "left": 0, "bottom": -1, "right": -1},
        # The complete used range preserves structural edits such as formats
        # and resized axes.  The content range is intentionally narrower and
        # drives the browser scrollbar so legacy blank-cell formatting cannot
        # make a small sheet appear thousands of rows long.
        "content_range": {"top": 0, "left": 0, "bottom": -1, "right": -1},
    }


def _is_meaningful_format(value: Any) -> bool:
    if not isinstance(value, dict):
        return False
    return any(item not in (None, False, "", "none", "default", {}, []) for item in value.values())


class SparseWorkbookStore:
    """Thread-safe, sparse V2 workbook state.

    It owns a new file.  If a V2 file has never been created, legacy values are
    copied into the first V2 sheet once; the legacy file is never modified.
    """

    def __init__(self, data_file: Path, legacy_file: Path | None = None) -> None:
        self.data_file = data_file
        self.legacy_file = legacy_file
        self.lock = threading.RLock()
        self.state = self._load()

    def _default_state(self) -> dict[str, Any]:
        return {
            "schema_version": 2,
            "revision": 0,
            "sheets": [_empty_sheet("sheet-1", "工作表1")],
        }

    def _load(self) -> dict[str, Any]:
        if self.data_file.exists():
            try:
                loaded = json.loads(self.data_file.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError) as exc:
                raise RuntimeError(f"Unable to load V2 workbook data: {self.data_file}") from exc
            if loaded.get("schema_version") != 2 or not isinstance(loaded.get("sheets"), list):
                raise RuntimeError("Unsupported V2 workbook state schema")
            self._normalize_state(loaded)
            return loaded
        state = self._migrate_legacy() or self._default_state()
        self._write(state)
        return state

    def _normalize_state(self, state: dict[str, Any]) -> None:
        state.setdefault("revision", 0)
        sheets = state.setdefault("sheets", [])
        if not sheets:
            sheets.append(_empty_sheet("sheet-1", "工作表1"))
        seen_ids: set[str] = set()
        for index, sheet in enumerate(sheets):
            identifier = str(sheet.get("id") or f"sheet-{index + 1}")
            while identifier in seen_ids:
                identifier = f"sheet-{uuid.uuid4().hex[:8]}"
            seen_ids.add(identifier)
            sheet["id"] = identifier
            sheet["name"] = str(sheet.get("name") or f"工作表{index + 1}")[:64]
            for key, default in (
                ("cells", {}),
                ("meta", {}),
                ("formats", {}),
                ("row_heights", {}),
                ("column_widths", {}),
                ("merges", []),
            ):
                sheet.setdefault(key, deepcopy(default))
            sheet.setdefault("used_range", {"top": 0, "left": 0, "bottom": -1, "right": -1})
            sheet.setdefault("content_range", {"top": 0, "left": 0, "bottom": -1, "right": -1})
            self._refresh_used_range(sheet)

    def _migrate_legacy(self) -> dict[str, Any] | None:
        if self.legacy_file is None or not self.legacy_file.exists():
            return None
        try:
            legacy = json.loads(self.legacy_file.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return None
        if legacy.get("schema_version") != 1 or not isinstance(legacy.get("rows"), list):
            return None
        state = self._default_state()
        sheet = state["sheets"][0]
        field_order = [
            "account_name", "teacher", "channel", "publish_count", "total_followers",
            "short_graphic_leads", "live_leads", "openings", "remark",
        ]
        field_order.extend(f"custom_{index}" for index in range(1, 92))
        field_positions = {field: index for index, field in enumerate(field_order)}
        row_positions = {int(row.get("id", index + 1)): index for index, row in enumerate(legacy["rows"])}
        for row_index, row in enumerate(legacy["rows"]):
            for col_index, field in enumerate(field_order):
                value = row.get(field)
                if value not in (None, ""):
                    sheet["cells"][coordinate_key(row_index, col_index)] = str(value)
        layout = legacy.get("layout") if isinstance(legacy.get("layout"), dict) else {}
        for field, width in (layout.get("column_widths") or {}).items():
            if field in field_positions and int(width) != DEFAULT_COLUMN_WIDTH:
                sheet["column_widths"][str(field_positions[field])] = int(width)
        for row_id, height in (layout.get("row_heights") or {}).items():
            position = row_positions.get(int(row_id))
            if position is not None and int(height) != DEFAULT_ROW_HEIGHT:
                sheet["row_heights"][str(position)] = int(height)
        for key, value in (layout.get("formats") or {}).items():
            try:
                row_id, field = key.split(":", 1)
                row_index = row_positions[int(row_id)]
                col_index = field_positions[field]
            except (KeyError, ValueError):
                continue
            if _is_meaningful_format(value):
                sheet["formats"][coordinate_key(row_index, col_index)] = deepcopy(value)
        for merge in layout.get("merges") or []:
            try:
                start_row = row_positions[int(merge["start_row_id"])]
                end_row = row_positions[int(merge["end_row_id"])]
                start_col = field_positions[str(merge["start_field"])]
                end_col = field_positions[str(merge["end_field"])]
            except (KeyError, TypeError, ValueError):
                continue
            sheet["merges"].append({
                "id": str(merge.get("id") or uuid.uuid4().hex),
                "start": {"row": min(start_row, end_row), "col": min(start_col, end_col)},
                "end": {"row": max(start_row, end_row), "col": max(start_col, end_col)},
            })
        self._refresh_used_range(sheet)
        return state

    def _write(self, state: dict[str, Any]) -> None:
        self.data_file.parent.mkdir(parents=True, exist_ok=True)
        descriptor, temporary_name = tempfile.mkstemp(dir=self.data_file.parent, prefix="workbook-v2-", suffix=".json.tmp")
        try:
            with os.fdopen(descriptor, "w", encoding="utf-8") as handle:
                json.dump(state, handle, ensure_ascii=False, indent=2)
                handle.flush()
                os.fsync(handle.fileno())
            os.replace(temporary_name, self.data_file)
        finally:
            if os.path.exists(temporary_name):
                os.unlink(temporary_name)

    def _sheet(self, sheet_id: str) -> dict[str, Any]:
        for sheet in self.state["sheets"]:
            if sheet["id"] == sheet_id:
                return sheet
        raise KeyError("找不到工作表")

    def _refresh_used_range(self, sheet: dict[str, Any]) -> dict[str, int]:
        content_points: list[tuple[int, int]] = []
        for key, value in sheet["cells"].items():
            if value not in (None, ""):
                content_points.append(parse_coordinate_key(key))
        if content_points:
            content_range = {
                "top": 0,
                "left": 0,
                "bottom": max(row for row, _ in content_points),
                "right": max(col for _, col in content_points),
            }
        else:
            content_range = {"top": 0, "left": 0, "bottom": -1, "right": -1}
        sheet["content_range"] = content_range

        points = list(content_points)
        for key, value in sheet["formats"].items():
            if _is_meaningful_format(value):
                points.append(parse_coordinate_key(key))
        for row in sheet["row_heights"]:
            points.append((int(row), 0))
        for col in sheet["column_widths"]:
            points.append((0, int(col)))
        for merge in sheet["merges"]:
            points.append((int(merge["start"]["row"]), int(merge["start"]["col"])))
            points.append((int(merge["end"]["row"]), int(merge["end"]["col"])))
        if points:
            used = {"top": 0, "left": 0, "bottom": max(row for row, _ in points), "right": max(col for _, col in points)}
        else:
            used = {"top": 0, "left": 0, "bottom": -1, "right": -1}
        sheet["used_range"] = used
        return used

    def sheets(self) -> list[dict[str, str]]:
        with self.lock:
            return [{"id": item["id"], "name": item["name"]} for item in self.state["sheets"]]

    def snapshot(self, sheet_id: str, top: int = 0, left: int = 0, rows: int = 120, cols: int = 60) -> dict[str, Any]:
        top = _coordinate(top, "起始行")
        left = _coordinate(left, "起始列")
        rows = _positive(rows, "窗口行数", 2000)
        cols = _positive(cols, "窗口列数", 2000)
        bottom = min(MAX_COORDINATE, top + rows - 1)
        right = min(MAX_COORDINATE, left + cols - 1)
        with self.lock:
            sheet = self._sheet(sheet_id)
            def in_window(key: str) -> bool:
                row, col = parse_coordinate_key(key)
                return top <= row <= bottom and left <= col <= right
            def axis_window(key: str, axis: str) -> bool:
                value = int(key)
                return (top <= value <= bottom) if axis == "row" else (left <= value <= right)
            merges = [
                deepcopy(merge)
                for merge in sheet["merges"]
                if not (
                    int(merge["end"]["row"]) < top
                    or int(merge["start"]["row"]) > bottom
                    or int(merge["end"]["col"]) < left
                    or int(merge["start"]["col"]) > right
                )
            ]
            return {
                "schema_version": 2,
                "revision": self.state["revision"],
                "sheets": self.sheets(),
                "sheet": {
                    "id": sheet["id"],
                    "name": sheet["name"],
                    "used_range": deepcopy(sheet["used_range"]),
                    "content_range": deepcopy(sheet["content_range"]),
                    "cells": {key: deepcopy(value) for key, value in sheet["cells"].items() if in_window(key)},
                    "meta": {key: deepcopy(value) for key, value in sheet["meta"].items() if in_window(key)},
                    "formats": {key: deepcopy(value) for key, value in sheet["formats"].items() if in_window(key)},
                    # Dimensions are sparse global axis metrics, not cells.
                    # Returning the override maps keeps every later window on
                    # the same coordinate system without materializing rows.
                    "row_heights": deepcopy(sheet["row_heights"]),
                    "column_widths": deepcopy(sheet["column_widths"]),
                    "merges": merges,
                },
                "window": {"top": top, "left": left, "bottom": bottom, "right": right},
            }

    def create_sheet(self, name: str) -> tuple[dict[str, Any], dict[str, Any]]:
        normalized = str(name).strip()[:64]
        if not normalized:
            raise ValueError("工作表名称不能为空")
        with self.lock:
            if any(sheet["name"] == normalized for sheet in self.state["sheets"]):
                raise ValueError("工作表名称已存在")
            sheet = _empty_sheet(f"sheet-{uuid.uuid4().hex[:12]}", normalized)
            self.state["sheets"].append(sheet)
            self.state["revision"] += 1
            self._write(self.state)
            return deepcopy(sheet), self._workbook_event("sheets")

    def rename_sheet(self, sheet_id: str, name: str) -> dict[str, Any]:
        normalized = str(name).strip()[:64]
        if not normalized:
            raise ValueError("工作表名称不能为空")
        with self.lock:
            if any(sheet["id"] != sheet_id and sheet["name"] == normalized for sheet in self.state["sheets"]):
                raise ValueError("工作表名称已存在")
            sheet = self._sheet(sheet_id)
            sheet["name"] = normalized
            self.state["revision"] += 1
            self._write(self.state)
            return self._workbook_event("sheets")

    def delete_sheet(self, sheet_id: str) -> dict[str, Any]:
        with self.lock:
            if len(self.state["sheets"]) <= 1:
                raise ValueError("至少保留一个工作表")
            self._sheet(sheet_id)
            self.state["sheets"] = [sheet for sheet in self.state["sheets"] if sheet["id"] != sheet_id]
            self.state["revision"] += 1
            self._write(self.state)
            return self._workbook_event("sheets")

    def _workbook_event(self, kind: str) -> dict[str, Any]:
        return {"kind": kind, "revision": self.state["revision"], "sheets": self.sheets()}

    def _validate_cells(self, cells: Iterable[Any]) -> list[dict[str, Any]]:
        normalized: list[dict[str, Any]] = []
        for item in cells:
            if not isinstance(item, dict):
                raise ValueError("单元格操作格式不正确")
            if len(normalized) >= MAX_BATCH_CELLS:
                raise ValueError("单次操作单元格过多")
            row = _coordinate(item.get("row"), "行号")
            col = _coordinate(item.get("col"), "列号")
            value = item.get("value")
            normalized.append({"row": row, "col": col, "value": "" if value is None else str(value), "version": int(item.get("version", 0))})
        return normalized

    def _apply_cells(self, sheet: dict[str, Any], cells: list[dict[str, Any]], name: str, client_id: str) -> list[dict[str, Any]]:
        events: list[dict[str, Any]] = []
        timestamp = utc_now()
        for item in cells:
            key = coordinate_key(item["row"], item["col"])
            current_meta = sheet["meta"].get(key, {})
            current_version = int(current_meta.get("version", 0))
            expected = item["version"]
            if expected and expected != current_version:
                raise ValueError("单元格已被其他协作者更新，请刷新后重试")
            value = item["value"]
            if value == "":
                sheet["cells"].pop(key, None)
            else:
                sheet["cells"][key] = value
            meta = {"version": current_version + 1, "updated_by": name, "updated_at": timestamp, "client_id": client_id}
            sheet["meta"][key] = meta
            events.append({"row": item["row"], "col": item["col"], "value": value, "meta": deepcopy(meta)})
        return events

    def _shift_axis(self, sheet: dict[str, Any], axis: str, at: int, count: int, deleting: bool) -> None:
        def remap_coordinate(row: int, col: int) -> tuple[int, int] | None:
            value = row if axis == "row" else col
            if deleting and at <= value < at + count:
                return None
            if value >= at + count and deleting:
                value -= count
            elif value >= at and not deleting:
                value += count
            return (value, col) if axis == "row" else (row, value)
        for name in ("cells", "meta", "formats"):
            source = sheet[name]
            target: dict[str, Any] = {}
            for key, value in source.items():
                mapped = remap_coordinate(*parse_coordinate_key(key))
                if mapped is not None:
                    target[coordinate_key(*mapped)] = value
            sheet[name] = target
        axis_key = "row_heights" if axis == "row" else "column_widths"
        target_axis: dict[str, Any] = {}
        for raw, value in sheet[axis_key].items():
            position = int(raw)
            if deleting and at <= position < at + count:
                continue
            if position >= at + count and deleting:
                position -= count
            elif position >= at and not deleting:
                position += count
            target_axis[str(position)] = value
        sheet[axis_key] = target_axis
        next_merges: list[dict[str, Any]] = []
        for merge in sheet["merges"]:
            start = int(merge["start"][axis])
            end = int(merge["end"][axis])
            if deleting:
                if end < at:
                    pass
                elif start >= at + count:
                    start -= count; end -= count
                else:
                    start = min(start, at)
                    end = max(start, end - count)
                    if start > end:
                        continue
            else:
                if start >= at:
                    start += count; end += count
                elif end >= at:
                    end += count
            merge["start"][axis] = start
            merge["end"][axis] = end
            next_merges.append(merge)
        sheet["merges"] = next_merges

    def apply_operations(self, sheet_id: str, operations: list[Any], name: str, client_id: str) -> dict[str, Any]:
        if not isinstance(operations, list) or not operations or len(operations) > MAX_BATCH_CELLS:
            raise ValueError("操作批次格式不正确")
        with self.lock:
            sheet = self._sheet(sheet_id)
            changed_cells: list[dict[str, Any]] = []
            changed_formats: list[dict[str, Any]] = []
            changed_rows: dict[str, int] = {}
            changed_cols: dict[str, int] = {}
            structure_changed = False
            for operation in operations:
                if not isinstance(operation, dict):
                    raise ValueError("操作格式不正确")
                kind = str(operation.get("type") or "")
                if kind == "cells":
                    changed_cells.extend(self._apply_cells(sheet, self._validate_cells(operation.get("cells") or []), name, client_id))
                elif kind == "range":
                    start = operation.get("start") or {}
                    top = _coordinate(start.get("row"), "起始行")
                    left = _coordinate(start.get("col"), "起始列")
                    values = operation.get("values")
                    if not isinstance(values, list):
                        raise ValueError("区域值格式不正确")
                    cells: list[dict[str, Any]] = []
                    for row_offset, row_values in enumerate(values):
                        if not isinstance(row_values, list):
                            raise ValueError("区域值格式不正确")
                        for col_offset, value in enumerate(row_values):
                            cells.append({"row": top + row_offset, "col": left + col_offset, "value": value})
                    changed_cells.extend(self._apply_cells(sheet, self._validate_cells(cells), name, client_id))
                elif kind == "format":
                    patch = operation.get("format")
                    if not isinstance(patch, dict):
                        raise ValueError("单元格格式不正确")
                    for item in self._validate_cells([{**cell, "value": ""} for cell in operation.get("cells") or []]):
                        key = coordinate_key(item["row"], item["col"])
                        current = dict(sheet["formats"].get(key, {}))
                        current.update(deepcopy(patch))
                        if _is_meaningful_format(current):
                            sheet["formats"][key] = current
                        else:
                            sheet["formats"].pop(key, None)
                        changed_formats.append({"row": item["row"], "col": item["col"], "format": deepcopy(sheet["formats"].get(key, {}))})
                elif kind == "layout":
                    action = str(operation.get("action") or "")
                    payload = operation.get("payload") or {}
                    if not isinstance(payload, dict):
                        raise ValueError("布局操作格式不正确")
                    if action == "row_height":
                        row = _coordinate(payload.get("row"), "行号")
                        height = int(payload.get("height", DEFAULT_ROW_HEIGHT))
                        if not 20 <= height <= 600:
                            raise ValueError("行高应在 20–600 像素之间")
                        if height == DEFAULT_ROW_HEIGHT:
                            sheet["row_heights"].pop(str(row), None)
                        else:
                            sheet["row_heights"][str(row)] = height
                        changed_rows[str(row)] = height
                    elif action == "column_width":
                        col = _coordinate(payload.get("col"), "列号")
                        width = int(payload.get("width", DEFAULT_COLUMN_WIDTH))
                        if not 40 <= width <= 1200:
                            raise ValueError("列宽应在 40–1200 像素之间")
                        if width == DEFAULT_COLUMN_WIDTH:
                            sheet["column_widths"].pop(str(col), None)
                        else:
                            sheet["column_widths"][str(col)] = width
                        changed_cols[str(col)] = width
                    elif action == "merge":
                        start = payload.get("start") or {}
                        end = payload.get("end") or {}
                        top = _coordinate(start.get("row"), "起始行"); left = _coordinate(start.get("col"), "起始列")
                        bottom = _coordinate(end.get("row"), "结束行"); right = _coordinate(end.get("col"), "结束列")
                        top, bottom = min(top, bottom), max(top, bottom)
                        left, right = min(left, right), max(left, right)
                        if top == bottom and left == right:
                            raise ValueError("请先选择两个或更多单元格")
                        for merge in sheet["merges"]:
                            overlaps = not (bottom < merge["start"]["row"] or top > merge["end"]["row"] or right < merge["start"]["col"] or left > merge["end"]["col"])
                            if overlaps:
                                raise ValueError("所选区域与已有合并单元格重叠")
                        sheet["merges"].append({"id": uuid.uuid4().hex, "start": {"row": top, "col": left}, "end": {"row": bottom, "col": right}})
                        structure_changed = True
                    elif action == "unmerge":
                        identifiers = {str(item) for item in payload.get("ids") or []}
                        sheet["merges"] = [merge for merge in sheet["merges"] if merge["id"] not in identifiers]
                        structure_changed = True
                    else:
                        raise ValueError("不支持的布局操作")
                elif kind in {"insert_rows", "delete_rows", "insert_columns", "delete_columns"}:
                    axis = "row" if kind.endswith("rows") else "col"
                    at = _coordinate(operation.get("at"), "起始位置")
                    count = _positive(operation.get("count", 1), "数量")
                    self._shift_axis(sheet, axis, at, count, kind.startswith("delete"))
                    structure_changed = True
                else:
                    raise ValueError("不支持的操作类型")
            self._refresh_used_range(sheet)
            self.state["revision"] += 1
            self._write(self.state)
            return {
                "kind": "delta",
                "sheet_id": sheet_id,
                "revision": self.state["revision"],
                "cells": changed_cells,
                "formats": changed_formats,
                "row_heights": changed_rows,
                "column_widths": changed_cols,
                "merges": deepcopy(sheet["merges"]) if structure_changed else None,
                "structure_changed": structure_changed,
                "used_range": deepcopy(sheet["used_range"]),
                "content_range": deepcopy(sheet["content_range"]),
            }
