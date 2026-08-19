"""Small, dependency-free collaboration service for GridBook.

The service owns its data and process. A standalone GridBook deployment
proxies the public /api prefix to this server.
"""

from __future__ import annotations

import gzip
import json
import os
import queue
import re
import tempfile
import threading
import time
import uuid
from copy import deepcopy
from datetime import datetime, timezone
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, unquote, urlparse

try:  # Support both `python backend/server.py` and package imports in tests.
    from .workbook_v2 import SparseWorkbookStore
except ImportError:  # pragma: no cover - script entry point
    from workbook_v2 import SparseWorkbookStore


HOST = os.environ.get("GRIDBOOK_HOST", "127.0.0.1")
PORT = int(os.environ.get("GRIDBOOK_PORT", "8767"))
DATA_FILE = Path(
    os.environ.get(
        "GRIDBOOK_DATA",
        str(Path(__file__).resolve().parent / "data" / "state.json"),
    )
)
V2_DATA_FILE = Path(
    os.environ.get(
        "GRIDBOOK_V2_DATA",
        str(DATA_FILE.with_name("workbook-v2.json")),
    )
)

TEXT_FIELDS = {
    "account_name": 64,
    "teacher": 32,
    "channel": 32,
    "remark": 240,
}
CUSTOM_FIELDS = {f"custom_{index}": 120 for index in range(1, 92)}
TEXT_FIELDS.update(CUSTOM_FIELDS)
NUMBER_FIELDS = {
    "publish_count",
    "total_followers",
    "short_graphic_leads",
    "live_leads",
    "openings",
}
EDITABLE_FIELDS = set(TEXT_FIELDS) | NUMBER_FIELDS
FIELD_ORDER = [
    "account_name",
    "teacher",
    "channel",
    "publish_count",
    "total_followers",
    "short_graphic_leads",
    "live_leads",
    "openings",
    "remark",
] + list(CUSTOM_FIELDS)
CELL_PATH = re.compile(r"^/api/cells/(?P<row_id>\d+)/(?P<field>[a-z_]+)$")
# Sparse range pastes and fill operations are sent as one transaction.  Keep
# the limit bounded, but large enough for a 1,000 × 20 spreadsheet paste.
MAX_BODY_BYTES = 8 * 1024 * 1024


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")


def blank_row(row_id: int) -> dict[str, Any]:
    """Return an empty spreadsheet row without inventing cell content."""
    row: dict[str, Any] = {
        "id": row_id,
        "account_name": "",
        "teacher": "",
        "channel": "",
        "publish_count": None,
        "total_followers": None,
        "short_graphic_leads": None,
        "live_leads": None,
        "openings": None,
        "remark": "",
    }
    row.update({key: "" for key in CUSTOM_FIELDS})
    return row


def default_state() -> dict[str, Any]:
    rows = [blank_row(row_id) for row_id in range(1, 201)]
    return {
        "schema_version": 1,
        "revision": 0,
        "next_row_id": 201,
        "period": {"start": "", "end": ""},
        "rows": rows,
        "cells": {},
        "layout": {
            "column_widths": {},
            "row_heights": {},
            "merges": [],
            "formats": {},
        },
        "activity": [],
    }


def validate_value(field: str, value: Any) -> str:
    """A blank spreadsheet has no field-level validation.

    The retained schema keys are used solely as sparse storage keys;
    every visible cell accepts arbitrary user text, including formulas as text.
    """
    if field not in TEXT_FIELDS and field not in NUMBER_FIELDS:
        raise ValueError("Unknown column")
    if value is None:
        return ""
    return value if isinstance(value, str) else str(value)


class CollaborationStore:
    def __init__(self, data_file: Path) -> None:
        self.data_file = data_file
        self.lock = threading.RLock()
        self.state = self._load()

    def _load(self) -> dict[str, Any]:
        if not self.data_file.exists():
            state = default_state()
            self._write(state)
            return state
        try:
            loaded = json.loads(self.data_file.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            raise RuntimeError(f"Unable to load GridBook data: {self.data_file}") from exc
        if loaded.get("schema_version") != 1 or not isinstance(loaded.get("rows"), list):
            raise RuntimeError("Unsupported spreadsheet state schema")
        loaded.setdefault(
            "layout",
            {"column_widths": {}, "row_heights": {}, "merges": [], "formats": {}},
        )
        loaded["layout"].setdefault("column_widths", {})
        loaded["layout"].setdefault("row_heights", {})
        loaded["layout"].setdefault("merges", [])
        loaded["layout"].setdefault("formats", {})
        return loaded

    def _write(self, state: dict[str, Any]) -> None:
        self.data_file.parent.mkdir(parents=True, exist_ok=True)
        file_descriptor, temporary_name = tempfile.mkstemp(
            dir=self.data_file.parent,
            prefix="state-",
            suffix=".json.tmp",
        )
        try:
            with os.fdopen(file_descriptor, "w", encoding="utf-8") as handle:
                json.dump(state, handle, ensure_ascii=False, indent=2)
                handle.flush()
                os.fsync(handle.fileno())
            os.replace(temporary_name, self.data_file)
        finally:
            if os.path.exists(temporary_name):
                os.unlink(temporary_name)

    def snapshot(self, limit: int | None = None) -> dict[str, Any]:
        with self.lock:
            if limit is None:
                return deepcopy(self.state)

            # The grid can grow to thousands of rows. A full deep copy
            # for every initial HTTP/SSE connection makes the first paint wait
            # on megabytes of JSON that are not visible yet.  Keep the same
            # state shape, but copy only the requested row window and the
            # metadata that belongs to it.
            safe_limit = max(1, min(int(limit), 10000))
            rows = deepcopy(self.state["rows"][:safe_limit])
            row_ids = {str(row["id"]) for row in rows}
            cells = {
                key: deepcopy(value)
                for key, value in self.state.get("cells", {}).items()
                if key.split(":", 1)[0] in row_ids
            }
            source_layout = self.state.get("layout", {})
            layout = {
                "column_widths": deepcopy(source_layout.get("column_widths", {})),
                "row_heights": {
                    key: value
                    for key, value in source_layout.get("row_heights", {}).items()
                    if str(key) in row_ids
                },
                "merges": deepcopy(source_layout.get("merges", [])),
                "formats": {
                    key: deepcopy(value)
                    for key, value in source_layout.get("formats", {}).items()
                    if key.split(":", 1)[0] in row_ids
                },
            }
            return {
                "schema_version": self.state.get("schema_version", 1),
                "revision": self.state.get("revision", 0),
                "next_row_id": self.state.get("next_row_id", 1),
                "period": deepcopy(self.state.get("period", {})),
                "rows": rows,
                "cells": cells,
                "layout": layout,
                "activity": deepcopy(self.state.get("activity", [])),
                "presence": [],
                "total_rows": len(self.state["rows"]),
            }

    def row_count(self) -> int:
        with self.lock:
            return len(self.state["rows"])

    def update_cell(
        self,
        row_id: int,
        field: str,
        value: Any,
        expected_version: int,
        name: str,
        client_id: str,
    ) -> dict[str, Any]:
        normalized = validate_value(field, value)
        with self.lock:
            row = next((item for item in self.state["rows"] if item["id"] == row_id), None)
            if row is None:
                raise KeyError("找不到该行")
            key = f"{row_id}:{field}"
            current_meta = self.state["cells"].get(key, {})
            current_version = int(current_meta.get("version", 0))
            if expected_version != current_version:
                return {
                    "conflict": True,
                    "row_id": row_id,
                    "field": field,
                    "value": row.get(field),
                    "meta": deepcopy(current_meta),
                }

            timestamp = utc_now()
            meta = {
                "version": current_version + 1,
                "updated_by": name,
                "updated_at": timestamp,
                "client_id": client_id,
            }
            row[field] = normalized
            self.state["cells"][key] = meta
            self.state["revision"] += 1
            activity = {
                "id": self.state["revision"],
                "kind": "cell",
                "row_id": row_id,
                "field": field,
                "field_label": FIELD_LABELS[field],
                "account_name": row.get("account_name", ""),
                "value": normalized,
                "name": name,
                "client_id": client_id,
                "updated_at": timestamp,
            }
            self.state["activity"] = [activity, *self.state["activity"]][:60]
            self._write(self.state)
            return {
                "conflict": False,
                "revision": self.state["revision"],
                "row_id": row_id,
                "field": field,
                "value": normalized,
                "meta": deepcopy(meta),
                "activity": deepcopy(activity),
            }

    def add_row(
        self,
        name: str,
        client_id: str,
        after_row_id: int | None = None,
    ) -> dict[str, Any]:
        with self.lock:
            row_id = int(self.state["next_row_id"])
            self.state["next_row_id"] = row_id + 1
            row = blank_row(row_id)
            timestamp = utc_now()
            if after_row_id is None:
                self.state["rows"].append(row)
            elif after_row_id == 0:
                self.state["rows"].insert(0, row)
            else:
                insert_at = next(
                    (index + 1 for index, item in enumerate(self.state["rows"]) if item["id"] == after_row_id),
                    len(self.state["rows"]),
                )
                self.state["rows"].insert(insert_at, row)
            self.state["revision"] += 1
            meta = {
                "version": 1,
                "updated_by": name,
                "updated_at": timestamp,
                "client_id": client_id,
            }
            activity = {
                "id": self.state["revision"],
                "kind": "row",
                "row_id": row_id,
                "field": "account_name",
                "field_label": "新增行",
                "account_name": "",
                "value": "",
                "name": name,
                "client_id": client_id,
                "updated_at": timestamp,
            }
            self.state["activity"] = [activity, *self.state["activity"]][:60]
            self._write(self.state)
            return {
                "revision": self.state["revision"],
                "row": deepcopy(row),
                "meta": deepcopy(meta),
                "activity": deepcopy(activity),
            }

    def ensure_rows(self, count: int, name: str, client_id: str) -> dict[str, Any]:
        if count < 1 or count > 10000:
            raise ValueError("工作表最多预留 10000 行")
        with self.lock:
            added: list[dict[str, Any]] = []
            while len(self.state["rows"]) < count:
                row_id = int(self.state["next_row_id"])
                self.state["next_row_id"] = row_id + 1
                row = blank_row(row_id)
                self.state["rows"].append(row)
                added.append(deepcopy(row))
            if added:
                self.state["revision"] += 1
                self._write(self.state)
            return {
                "revision": self.state["revision"],
                "added": added,
                "rows": len(self.state["rows"]),
                "name": name,
                "client_id": client_id,
            }

    def delete_rows(self, row_ids: list[int], name: str, client_id: str) -> dict[str, Any]:
        normalized = list(dict.fromkeys(int(row_id) for row_id in row_ids))
        if not normalized or len(normalized) > 100:
            raise ValueError("一次可删除 1–100 行")
        with self.lock:
            existing = {int(row["id"]) for row in self.state["rows"]}
            targets = [row_id for row_id in normalized if row_id in existing]
            if not targets:
                raise ValueError("找不到要删除的行")
            if len(targets) >= len(self.state["rows"]):
                raise ValueError("至少需要保留一行")
            target_set = set(targets)
            self.state["rows"] = [row for row in self.state["rows"] if int(row["id"]) not in target_set]
            self.state["cells"] = {
                key: value
                for key, value in self.state["cells"].items()
                if int(key.split(":", 1)[0]) not in target_set
            }
            layout = self.state["layout"]
            layout["row_heights"] = {
                key: value for key, value in layout["row_heights"].items() if int(key) not in target_set
            }
            layout["formats"] = {
                key: value
                for key, value in layout["formats"].items()
                if int(key.split(":", 1)[0]) not in target_set
            }
            layout["merges"] = [
                merge
                for merge in layout["merges"]
                if int(merge["start_row_id"]) not in target_set
                and int(merge["end_row_id"]) not in target_set
            ]
            self.state["revision"] += 1
            self._write(self.state)
            return {
                "revision": self.state["revision"],
                "row_ids": targets,
                "layout": deepcopy(layout),
                "name": name,
                "client_id": client_id,
            }

    def update_layout(
        self,
        action: str,
        payload: dict[str, Any],
        name: str,
        client_id: str,
    ) -> dict[str, Any]:
        with self.lock:
            layout = self.state["layout"]
            rows = self.state["rows"]
            row_ids = [int(row["id"]) for row in rows]
            row_positions = {row_id: index for index, row_id in enumerate(row_ids)}

            if action == "column_width":
                field = str(payload.get("field") or "")
                width = int(payload.get("width", 0))
                if field not in EDITABLE_FIELDS or not 64 <= width <= 520:
                    raise ValueError("列宽应在 64–520 像素之间")
                layout["column_widths"][field] = width
            elif action == "row_height":
                row_id = int(payload.get("row_id", 0))
                height = int(payload.get("height", 0))
                if row_id not in row_positions or not 30 <= height <= 240:
                    raise ValueError("行高应在 30–240 像素之间")
                layout["row_heights"][str(row_id)] = height
            elif action == "merge":
                start_row_id = int(payload.get("start_row_id", 0))
                end_row_id = int(payload.get("end_row_id", 0))
                start_field = str(payload.get("start_field") or "")
                end_field = str(payload.get("end_field") or "")
                if (
                    start_row_id not in row_positions
                    or end_row_id not in row_positions
                    or start_field not in EDITABLE_FIELDS
                    or end_field not in EDITABLE_FIELDS
                ):
                    raise ValueError("合并区域无效")
                top = min(row_positions[start_row_id], row_positions[end_row_id])
                bottom = max(row_positions[start_row_id], row_positions[end_row_id])
                left = min(FIELD_ORDER.index(start_field), FIELD_ORDER.index(end_field))
                right = max(FIELD_ORDER.index(start_field), FIELD_ORDER.index(end_field))
                if top == bottom and left == right:
                    raise ValueError("请先选择两个或更多单元格")
                for merge in layout["merges"]:
                    merge_top = row_positions.get(int(merge["start_row_id"]), -1)
                    merge_bottom = row_positions.get(int(merge["end_row_id"]), -1)
                    merge_left = FIELD_ORDER.index(merge["start_field"])
                    merge_right = FIELD_ORDER.index(merge["end_field"])
                    overlaps = not (
                        bottom < merge_top
                        or top > merge_bottom
                        or right < merge_left
                        or left > merge_right
                    )
                    if overlaps:
                        raise ValueError("所选区域与已有合并单元格重叠")
                layout["merges"].append(
                    {
                        "id": uuid.uuid4().hex,
                        "start_row_id": row_ids[top],
                        "end_row_id": row_ids[bottom],
                        "start_field": FIELD_ORDER[left],
                        "end_field": FIELD_ORDER[right],
                    }
                )
            elif action == "unmerge":
                merge_ids = {str(item) for item in payload.get("merge_ids", [])}
                layout["merges"] = [item for item in layout["merges"] if item["id"] not in merge_ids]
            elif action == "format":
                cells = payload.get("cells")
                patch = payload.get("format")
                replace_format = payload.get("replace") is True
                border_action = payload.get("border_action")
                border_spec = payload.get("border")
                if border_action is not None:
                    if border_action not in {"none", "all", "outside", "top", "bottom", "left", "right"}:
                        raise ValueError("Unsupported border action")
                    if border_spec is None:
                        border_spec = {"style": "solid", "width": 1, "color": "#66717d"}
                if not isinstance(cells, list) or not cells or len(cells) > 100_000 or not isinstance(patch, dict):
                    raise ValueError("格式化区域无效")
                allowed: dict[str, set[Any]] = {
                    "bold": {True, False},
                    "italic": {True, False},
                    "underline": {True, False},
                    "border": {True, False},
                    "wrap": {True, False},
                    "align": {"left", "center", "right", "justify"},
                    "vertical_align": {"top", "middle", "bottom"},
                    "fill": {"none", "yellow", "blue", "green", "rose", "#fff2cc", "#ddebf7", "#e2f0d9", "#fce4d6", "#f4cccc", "#d9e1f2", "#c6e0b4", "#ffd966", "#5b9bd5", "#70ad47", "#ed7d31", "#c00000", "#203864"},
                    "font_color": {"default", "red", "blue", "green", "#000000", "#ffffff", "#7f7f7f", "#c00000", "#ff0000", "#ffc000", "#ffff00", "#92d050", "#00b050", "#00b0f0", "#0070c0", "#002060", "#7030a0"},
                    "font_family": {"Arial", "Microsoft YaHei", "SimSun", "FangSong", "仿宋_GB2312", "KaiTi", "楷体_GB2312", "SimHei", "方正小标宋简体", "方正大标宋简体", "方正仿宋简体", "方正仿宋_GBK", "方正楷体简体", "方正楷体_GBK", "方正黑体简体", "方正黑体_GBK", "Times New Roman"},
                    "number_format": {"general", "number", "date", "percentage", "text"},
                }
                def is_hex_color(value: Any) -> bool:
                    return isinstance(value, str) and bool(re.fullmatch(r"#[0-9a-fA-F]{3}(?:[0-9a-fA-F]{3})?", value))

                normalized_patch: dict[str, Any] = {}
                for key, value in patch.items():
                    if key == "borders":
                        if not isinstance(value, dict):
                            raise ValueError("Unsupported border format")
                        normalized_patch[key] = value
                        continue
                    if key == "font_size":
                        try:
                            numeric_size = float(value)
                        except (TypeError, ValueError):
                            numeric_size = -1
                        if isinstance(value, bool) or not numeric_size.is_integer() or not 8 <= numeric_size <= 72:
                            raise ValueError("Unsupported cell format")
                        normalized_patch[key] = int(numeric_size)
                        continue
                    if key not in allowed:
                        raise ValueError("不支持的单元格格式")
                    if key in {"fill", "font_color"} and is_hex_color(value):
                        normalized_patch[key] = value.lower()
                        continue
                    if value not in allowed[key]:
                        raise ValueError("不支持的单元格格式")
                    normalized_patch[key] = value
                validated_cells: list[tuple[int, str, int, int]] = []
                for cell in cells:
                    row_id = int(cell.get("row_id", 0))
                    field = str(cell.get("field") or "")
                    if row_id not in row_positions or field not in EDITABLE_FIELDS:
                        raise ValueError("格式化区域包含无效单元格")
                    validated_cells.append((row_id, field, row_positions[row_id], FIELD_ORDER.index(field)))

                if border_action is not None:
                    min_row = min(item[2] for item in validated_cells)
                    max_row = max(item[2] for item in validated_cells)
                    min_col = min(item[3] for item in validated_cells)
                    max_col = max(item[3] for item in validated_cells)
                    for row_id, field, row_index, col_index in validated_cells:
                        key = f"{row_id}:{field}"
                        current = dict(layout["formats"].get(key, {}))
                        borders = dict(current.get("borders") or {})
                        if current.get("border") is True and not borders:
                            borders = {side: dict(border_spec) for side in ("top", "right", "bottom", "left")}
                        if border_action == "none":
                            borders = {}
                        else:
                            sides: set[str] = set()
                            if border_action == "all":
                                sides = {"top", "right", "bottom", "left"}
                            elif border_action == "outside":
                                if row_index == min_row: sides.add("top")
                                if row_index == max_row: sides.add("bottom")
                                if col_index == min_col: sides.add("left")
                                if col_index == max_col: sides.add("right")
                            else:
                                sides = {str(border_action)}
                            for side in sides:
                                borders[side] = border_spec
                        if borders:
                            current["borders"] = borders
                            current["border"] = True
                        else:
                            current.pop("borders", None)
                            current["border"] = False
                        layout["formats"][key] = current
                    normalized_patch = {}

                for row_id, field, _row_index, _col_index in validated_cells:
                    key = f"{row_id}:{field}"
                    current = {} if replace_format else dict(layout["formats"].get(key, {}))
                    current.update(normalized_patch)
                    layout["formats"][key] = current
            else:
                raise ValueError("不支持的布局操作")

            self.state["revision"] += 1
            self._write(self.state)
            return {
                "revision": self.state["revision"],
                "layout": deepcopy(layout),
                "name": name,
                "client_id": client_id,
            }

    def batch_update_cells(
        self,
        updates: list[dict[str, Any]],
        name: str,
        client_id: str,
    ) -> list[dict[str, Any]]:
        if not updates or len(updates) > 100_000:
            raise ValueError("一次操作最多 100000 个单元格")

        prepared: list[tuple[dict[str, Any], str, str | int | None, int]] = []
        seen: set[str] = set()
        with self.lock:
            rows_by_id = {int(row["id"]): row for row in self.state["rows"]}
            conflicts: list[dict[str, Any]] = []
            for update in updates:
                try:
                    row_id = int(update.get("row_id"))
                    expected_version = int(update.get("version", 0))
                except (TypeError, ValueError) as exc:
                    raise ValueError("粘贴区域包含无效位置") from exc
                field = str(update.get("field") or "")
                row = rows_by_id.get(row_id)
                if row is None:
                    raise ValueError("粘贴区域超出了当前表格")
                if field not in EDITABLE_FIELDS:
                    raise ValueError("粘贴区域包含不可编辑字段")
                key = f"{row_id}:{field}"
                if key in seen:
                    raise ValueError("粘贴区域包含重复单元格")
                seen.add(key)
                normalized = validate_value(field, update.get("value"))
                current_meta = self.state["cells"].get(key, {})
                current_version = int(current_meta.get("version", 0))
                if current_version != expected_version:
                    conflicts.append(
                        {
                            "row_id": row_id,
                            "field": field,
                            "value": row.get(field),
                            "meta": deepcopy(current_meta),
                        }
                    )
                prepared.append((row, field, normalized, current_version))

            if conflicts:
                return [{"conflict": True, "conflicts": conflicts}]

            events: list[dict[str, Any]] = []
            for row, field, normalized, current_version in prepared:
                row_id = int(row["id"])
                timestamp = utc_now()
                meta = {
                    "version": current_version + 1,
                    "updated_by": name,
                    "updated_at": timestamp,
                    "client_id": client_id,
                }
                row[field] = normalized
                self.state["cells"][f"{row_id}:{field}"] = meta
                self.state["revision"] += 1
                activity = {
                    "id": self.state["revision"],
                    "kind": "cell",
                    "row_id": row_id,
                    "field": field,
                    "field_label": FIELD_LABELS[field],
                    "account_name": row.get("account_name", ""),
                    "value": normalized,
                    "name": name,
                    "client_id": client_id,
                    "updated_at": timestamp,
                }
                self.state["activity"].insert(0, activity)
                events.append(
                    {
                        "conflict": False,
                        "revision": self.state["revision"],
                        "row_id": row_id,
                        "field": field,
                        "value": normalized,
                        "meta": deepcopy(meta),
                        "activity": deepcopy(activity),
                    }
                )
            self.state["activity"] = self.state["activity"][:60]
            self._write(self.state)
            return events

    def apply_operations(
        self,
        operations: list[dict[str, Any]],
        name: str,
        client_id: str,
        base_revision: int,
    ) -> list[dict[str, Any]]:
        """Apply a compact operation batch while keeping legacy endpoints intact."""
        if base_revision < 0:
            raise ValueError("Invalid base revision")
        cell_updates: list[dict[str, Any]] = []
        layout_events: list[dict[str, Any]] = []
        fill_format_groups: dict[str, dict[str, Any]] = {}
        for operation in operations:
            kind = str(operation.get("type") or "")
            if kind == "cell":
                cell_updates.append(operation)
            elif kind == "range":
                start = operation.get("start") or {}
                values = operation.get("values") or []
                for row_offset, row_values in enumerate(values):
                    for col_offset, value in enumerate(row_values):
                        cell_updates.append({
                            "row_id": int(start.get("row_id", 0)) + row_offset,
                            "field": str(operation.get("fields", [])[col_offset]) if operation.get("fields") else "",
                            "value": value,
                            "version": 0,
                        })
            elif kind == "fill":
                updates = operation.get("updates") or []
                formats = operation.get("formats") or []
                if not isinstance(updates, list) or len(updates) > 100_000:
                    raise ValueError("填充操作包含无效单元格")
                if not isinstance(formats, list) or len(formats) > 100_000:
                    raise ValueError("填充格式包含无效单元格")
                cell_updates.extend(updates)
                for item in formats:
                    if not isinstance(item, dict):
                        raise ValueError("填充格式格式不正确")
                    row_id = int(item.get("row_id", 0))
                    field = str(item.get("field") or "")
                    patch = item.get("format")
                    if not isinstance(patch, dict):
                        raise ValueError("填充格式格式不正确")
                    key = json.dumps(patch, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
                    group = fill_format_groups.setdefault(key, {"format": patch, "cells": []})
                    group["cells"].append({"row_id": row_id, "field": field})
            elif kind == "layout":
                action = str(operation.get("action") or "")
                layout_events.append(self.update_layout(action, dict(operation.get("payload") or {}), name, client_id))
            else:
                raise ValueError("Unsupported operation type")
        events: list[dict[str, Any]] = []
        if cell_updates:
            events.extend(self.batch_update_cells(cell_updates, name, client_id))
            if events and events[0].get("conflict"):
                return events
        for group in fill_format_groups.values():
            layout_events.append(
                self.update_layout(
                    "format",
                    {"cells": group["cells"], "format": group["format"], "replace": True},
                    name,
                    client_id,
                )
            )
        events.extend(layout_events)
        return events


def excel_column_name(index: int) -> str:
    value = index + 1
    label = ""
    while value > 0:
        value, remainder = divmod(value - 1, 26)
        label = chr(65 + remainder) + label
    return label


FIELD_LABELS = {field: excel_column_name(index) for index, field in enumerate(FIELD_ORDER)}


STORE = CollaborationStore(DATA_FILE)
V2_STORE = SparseWorkbookStore(V2_DATA_FILE, DATA_FILE)
CLIENTS: dict[str, tuple[queue.Queue[tuple[str, dict[str, Any]]], str]] = {}
CLIENTS_LOCK = threading.RLock()
PRESENCE: dict[str, dict[str, Any]] = {}
PRESENCE_LOCK = threading.RLock()
V2_CLIENTS: dict[str, tuple[queue.Queue[tuple[str, dict[str, Any]]], str, str]] = {}
V2_CLIENTS_LOCK = threading.RLock()
V2_PRESENCE: dict[str, dict[str, Any]] = {}
V2_PRESENCE_LOCK = threading.RLock()


def normalize_identity(client_id: Any, name: Any) -> tuple[str, str]:
    normalized_client = str(client_id or "").strip()
    normalized_name = str(name or "").strip()
    if not re.fullmatch(r"[a-zA-Z0-9-]{8,80}", normalized_client):
        raise ValueError("无效的协作会话")
    if not normalized_name or len(normalized_name) > 24:
        raise ValueError("显示名称应为 1–24 个字符")
    return normalized_client, normalized_name


def presence_snapshot() -> list[dict[str, Any]]:
    threshold = time.time() - 50
    with PRESENCE_LOCK:
        expired = [key for key, item in PRESENCE.items() if item["last_seen_epoch"] < threshold]
        for key in expired:
            PRESENCE.pop(key, None)
        return [
            {key: value for key, value in item.items() if key != "last_seen_epoch"}
            for item in PRESENCE.values()
        ]


def touch_presence(client_id: str, name: str, cell: Any = None) -> None:
    safe_cell = None
    if isinstance(cell, dict):
        try:
            row_id = int(cell.get("row_id"))
        except (TypeError, ValueError):
            row_id = 0
        field = str(cell.get("field") or "")
        if row_id > 0 and field in EDITABLE_FIELDS:
            safe_cell = {"row_id": row_id, "field": field}
    with PRESENCE_LOCK:
        PRESENCE[client_id] = {
            "client_id": client_id,
            "name": name,
            "cell": safe_cell,
            "last_seen": utc_now(),
            "last_seen_epoch": time.time(),
        }


def remove_presence_if_disconnected(client_id: str) -> None:
    with CLIENTS_LOCK:
        still_connected = any(item[1] == client_id for item in CLIENTS.values())
    if not still_connected:
        with PRESENCE_LOCK:
            PRESENCE.pop(client_id, None)


def broadcast(event: str, payload: dict[str, Any]) -> None:
    with CLIENTS_LOCK:
        clients = list(CLIENTS.values())
    for event_queue, _ in clients:
        try:
            event_queue.put_nowait((event, payload))
        except queue.Full:
            try:
                event_queue.get_nowait()
                event_queue.put_nowait((event, payload))
            except (queue.Empty, queue.Full):
                pass


def v2_presence_snapshot(sheet_id: str) -> list[dict[str, Any]]:
    threshold = time.time() - 50
    with V2_PRESENCE_LOCK:
        expired = [key for key, item in V2_PRESENCE.items() if item["last_seen_epoch"] < threshold]
        for key in expired:
            V2_PRESENCE.pop(key, None)
        return [
            {key: value for key, value in item.items() if key != "last_seen_epoch"}
            for item in V2_PRESENCE.values()
            if item["sheet_id"] == sheet_id
        ]


def touch_v2_presence(client_id: str, name: str, sheet_id: str, cell: Any = None) -> None:
    safe_cell = None
    if isinstance(cell, dict):
        try:
            row = int(cell.get("row"))
            col = int(cell.get("col"))
        except (TypeError, ValueError):
            row, col = -1, -1
        if row >= 0 and col >= 0:
            safe_cell = {"row": row, "col": col}
    with V2_PRESENCE_LOCK:
        V2_PRESENCE[client_id] = {
            "client_id": client_id,
            "name": name,
            "sheet_id": sheet_id,
            "cell": safe_cell,
            "last_seen": utc_now(),
            "last_seen_epoch": time.time(),
        }


def v2_broadcast(sheet_id: str, event: str, payload: dict[str, Any]) -> None:
    with V2_CLIENTS_LOCK:
        clients = [item for item in V2_CLIENTS.values() if item[2] == sheet_id]
    for event_queue, _, _ in clients:
        try:
            event_queue.put_nowait((event, payload))
        except queue.Full:
            try:
                event_queue.get_nowait()
                event_queue.put_nowait((event, payload))
            except (queue.Empty, queue.Full):
                pass


def v2_broadcast_all(event: str, payload: dict[str, Any]) -> None:
    with V2_CLIENTS_LOCK:
        clients = list(V2_CLIENTS.values())
    for event_queue, _, _ in clients:
        try:
            event_queue.put_nowait((event, payload))
        except queue.Full:
            try:
                event_queue.get_nowait()
                event_queue.put_nowait((event, payload))
            except (queue.Empty, queue.Full):
                pass


class GridBookHandler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"
    server_version = "GridBook/0.1"

    def log_message(self, message: str, *args: Any) -> None:
        print(f"{self.address_string()} - {message % args}", flush=True)

    def _json_body(self) -> dict[str, Any]:
        content_type = self.headers.get("Content-Type", "").split(";", 1)[0].strip()
        if content_type != "application/json":
            raise ValueError("请求需要使用 application/json")
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError as exc:
            raise ValueError("无效的请求长度") from exc
        if length <= 0 or length > MAX_BODY_BYTES:
            raise ValueError("请求内容为空或过大")
        try:
            payload = json.loads(self.rfile.read(length))
        except json.JSONDecodeError as exc:
            raise ValueError("JSON 格式不正确") from exc
        if not isinstance(payload, dict):
            raise ValueError("请求内容必须为对象")
        return payload

    def _send_json(self, status: int, payload: dict[str, Any]) -> None:
        encoded = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        compressed = gzip.compress(encoded, compresslevel=6)
        use_compressed = len(compressed) < len(encoded)
        body = compressed if use_compressed else encoded
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        if use_compressed:
            self.send_header("Content-Encoding", "gzip")
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.end_headers()
        self.wfile.write(body)

    def _send_error_json(self, status: int, message: str) -> None:
        self._send_json(status, {"error": message})

    def _send_sse(self, event: str, payload: dict[str, Any]) -> None:
        encoded = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
        message = f"event: {event}\ndata: {encoded}\n\n".encode("utf-8")
        self.wfile.write(message)
        self.wfile.flush()

    def _v2_window(self, parsed: Any) -> None:
        prefix = "/api/v2/sheets/"
        suffix = "/window"
        sheet_id = unquote(parsed.path[len(prefix):-len(suffix)])
        query = parse_qs(parsed.query)
        try:
            snapshot = V2_STORE.snapshot(
                sheet_id,
                int(query.get("top", [0])[0]),
                int(query.get("left", [0])[0]),
                int(query.get("rows", [120])[0]),
                int(query.get("cols", [60])[0]),
            )
        except KeyError as exc:
            self._send_error_json(HTTPStatus.NOT_FOUND, str(exc.args[0]))
            return
        except (TypeError, ValueError) as exc:
            self._send_error_json(HTTPStatus.BAD_REQUEST, str(exc))
            return
        snapshot["presence"] = v2_presence_snapshot(sheet_id)
        self._send_json(HTTPStatus.OK, snapshot)

    def _v2_events(self, raw_query: str) -> None:
        query = parse_qs(raw_query)
        try:
            client_id, name = normalize_identity(query.get("client_id", [""])[0], query.get("name", [""])[0])
            sheet_id = str(query.get("sheet_id", [""])[0])
            top = int(query.get("top", [0])[0])
            left = int(query.get("left", [0])[0])
            rows = int(query.get("rows", [120])[0])
            cols = int(query.get("cols", [60])[0])
            V2_STORE.snapshot(sheet_id, top, left, rows, cols)
        except KeyError as exc:
            self._send_error_json(HTTPStatus.NOT_FOUND, str(exc.args[0]))
            return
        except (TypeError, ValueError) as exc:
            self._send_error_json(HTTPStatus.BAD_REQUEST, str(exc))
            return

        connection_id = uuid.uuid4().hex
        event_queue: queue.Queue[tuple[str, dict[str, Any]]] = queue.Queue(maxsize=200)
        with V2_CLIENTS_LOCK:
            V2_CLIENTS[connection_id] = (event_queue, client_id, sheet_id)
        touch_v2_presence(client_id, name, sheet_id)

        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "text/event-stream; charset=utf-8")
        self.send_header("Cache-Control", "no-cache, no-transform")
        self.send_header("Connection", "keep-alive")
        self.send_header("X-Accel-Buffering", "no")
        self.end_headers()
        try:
            snapshot = V2_STORE.snapshot(sheet_id, top, left, rows, cols)
            snapshot["presence"] = v2_presence_snapshot(sheet_id)
            self._send_sse("snapshot", snapshot)
            v2_broadcast(sheet_id, "presence", {"sheet_id": sheet_id, "presence": v2_presence_snapshot(sheet_id)})
            while True:
                try:
                    event, payload = event_queue.get(timeout=15)
                    self._send_sse(event, payload)
                except queue.Empty:
                    self.wfile.write(b": keep-alive\n\n")
                    self.wfile.flush()
        except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError, TimeoutError, OSError):
            pass
        finally:
            with V2_CLIENTS_LOCK:
                V2_CLIENTS.pop(connection_id, None)
                still_connected = any(item[1] == client_id for item in V2_CLIENTS.values())
            if not still_connected:
                with V2_PRESENCE_LOCK:
                    V2_PRESENCE.pop(client_id, None)
            v2_broadcast(sheet_id, "presence", {"sheet_id": sheet_id, "presence": v2_presence_snapshot(sheet_id)})

    def _v2_post(self, path: str) -> None:
        try:
            payload = self._json_body()
            client_id, name = normalize_identity(payload.get("client_id"), payload.get("name"))
        except ValueError as exc:
            self._send_error_json(HTTPStatus.BAD_REQUEST, str(exc))
            return

        if path == "/api/v2/presence":
            sheet_id = str(payload.get("sheet_id") or "")
            try:
                V2_STORE.snapshot(sheet_id, 0, 0, 1, 1)
            except KeyError as exc:
                self._send_error_json(HTTPStatus.NOT_FOUND, str(exc.args[0]))
                return
            except ValueError as exc:
                self._send_error_json(HTTPStatus.BAD_REQUEST, str(exc))
                return
            touch_v2_presence(client_id, name, sheet_id, payload.get("cell"))
            present = v2_presence_snapshot(sheet_id)
            v2_broadcast(sheet_id, "presence", {"sheet_id": sheet_id, "presence": present})
            self._send_json(HTTPStatus.OK, {"presence": present})
            return

        if path == "/api/v2/sheets":
            if str(payload.get("action") or "create") != "create":
                self._send_error_json(HTTPStatus.BAD_REQUEST, "不支持的工作表操作")
                return
            try:
                sheet, event = V2_STORE.create_sheet(str(payload.get("sheet_name") or ""))
            except ValueError as exc:
                self._send_error_json(HTTPStatus.BAD_REQUEST, str(exc))
                return
            v2_broadcast_all("workbook", event)
            self._send_json(HTTPStatus.CREATED, {"sheet": sheet, **event})
            return

        match = re.fullmatch(r"/api/v2/sheets/([^/]+)/ops", path)
        if match:
            sheet_id = unquote(match.group(1))
            try:
                result = V2_STORE.apply_operations(sheet_id, payload.get("operations"), name, client_id)
            except KeyError as exc:
                self._send_error_json(HTTPStatus.NOT_FOUND, str(exc.args[0]))
                return
            except (TypeError, ValueError) as exc:
                self._send_error_json(HTTPStatus.BAD_REQUEST, str(exc))
                return
            touch_v2_presence(client_id, name, sheet_id, payload.get("cell"))
            v2_broadcast(sheet_id, "delta", result)
            v2_broadcast(sheet_id, "presence", {"sheet_id": sheet_id, "presence": v2_presence_snapshot(sheet_id)})
            self._send_json(HTTPStatus.OK, result)
            return

        match = re.fullmatch(r"/api/v2/sheets/([^/]+)/rename", path)
        if match:
            sheet_id = unquote(match.group(1))
            try:
                event = V2_STORE.rename_sheet(sheet_id, str(payload.get("sheet_name") or ""))
            except KeyError as exc:
                self._send_error_json(HTTPStatus.NOT_FOUND, str(exc.args[0]))
                return
            except ValueError as exc:
                self._send_error_json(HTTPStatus.BAD_REQUEST, str(exc))
                return
            v2_broadcast_all("workbook", event)
            self._send_json(HTTPStatus.OK, event)
            return

        self._send_error_json(HTTPStatus.NOT_FOUND, "接口不存在")

    def _v2_delete(self, path: str) -> None:
        match = re.fullmatch(r"/api/v2/sheets/([^/]+)", path)
        if not match:
            self._send_error_json(HTTPStatus.NOT_FOUND, "接口不存在")
            return
        try:
            payload = self._json_body()
            normalize_identity(payload.get("client_id"), payload.get("name"))
            event = V2_STORE.delete_sheet(unquote(match.group(1)))
        except KeyError as exc:
            self._send_error_json(HTTPStatus.NOT_FOUND, str(exc.args[0]))
            return
        except (TypeError, ValueError) as exc:
            self._send_error_json(HTTPStatus.BAD_REQUEST, str(exc))
            return
        v2_broadcast_all("workbook", event)
        self._send_json(HTTPStatus.OK, event)

    def do_GET(self) -> None:  # noqa: N802
        parsed = urlparse(self.path)
        if parsed.path == "/api/v2/sheets":
            self._send_json(HTTPStatus.OK, {"revision": V2_STORE.state["revision"], "sheets": V2_STORE.sheets()})
            return
        if parsed.path.startswith("/api/v2/sheets/") and parsed.path.endswith("/window"):
            self._v2_window(parsed)
            return
        if parsed.path == "/api/v2/events":
            self._v2_events(parsed.query)
            return
        if parsed.path == "/api/health":
            self._send_json(HTTPStatus.OK, {"ok": True, "service": "gridbook"})
            return
        if parsed.path == "/api/state":
            query = parse_qs(parsed.query)
            raw_limit = query.get("limit", [""])[0]
            try:
                limit = max(1, min(int(raw_limit), 10000)) if raw_limit else None
            except ValueError:
                limit = None
            snapshot = STORE.snapshot(limit)
            snapshot["presence"] = presence_snapshot()
            self._send_json(HTTPStatus.OK, snapshot)
            return
        if parsed.path == "/api/events":
            self._events(parsed.query)
            return
        self._send_error_json(HTTPStatus.NOT_FOUND, "接口不存在")

    def _events(self, raw_query: str) -> None:
        query = parse_qs(raw_query)
        try:
            client_id, name = normalize_identity(
                query.get("client_id", [""])[0],
                query.get("name", [""])[0],
            )
        except ValueError as exc:
            self._send_error_json(HTTPStatus.BAD_REQUEST, str(exc))
            return

        connection_id = uuid.uuid4().hex
        event_queue: queue.Queue[tuple[str, dict[str, Any]]] = queue.Queue(maxsize=100)
        with CLIENTS_LOCK:
            CLIENTS[connection_id] = (event_queue, client_id)
        touch_presence(client_id, name)

        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "text/event-stream; charset=utf-8")
        self.send_header("Cache-Control", "no-cache, no-transform")
        self.send_header("Connection", "keep-alive")
        self.send_header("X-Accel-Buffering", "no")
        self.end_headers()
        try:
            query_limit = parse_qs(raw_query).get("limit", [""])[0]
            try:
                limit = max(1, min(int(query_limit), 10000)) if query_limit else None
            except ValueError:
                limit = None
            snapshot = STORE.snapshot(limit)
            snapshot["presence"] = presence_snapshot()
            self._send_sse("snapshot", snapshot)
            broadcast("presence", {"presence": presence_snapshot()})
            while True:
                try:
                    event, payload = event_queue.get(timeout=15)
                    self._send_sse(event, payload)
                except queue.Empty:
                    self.wfile.write(b": keep-alive\n\n")
                    self.wfile.flush()
        except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError, TimeoutError, OSError):
            pass
        finally:
            with CLIENTS_LOCK:
                CLIENTS.pop(connection_id, None)
            remove_presence_if_disconnected(client_id)
            broadcast("presence", {"presence": presence_snapshot()})

    def do_PATCH(self) -> None:  # noqa: N802
        path = urlparse(self.path).path
        match = CELL_PATH.fullmatch(path)
        if match is None:
            self._send_error_json(HTTPStatus.NOT_FOUND, "接口不存在")
            return
        field = unquote(match.group("field"))
        if field not in EDITABLE_FIELDS:
            self._send_error_json(HTTPStatus.BAD_REQUEST, "该字段不可编辑")
            return
        try:
            payload = self._json_body()
            client_id, name = normalize_identity(payload.get("client_id"), payload.get("name"))
            expected_version = int(payload.get("version", 0))
            if expected_version < 0:
                raise ValueError("无效的单元格版本")
            result = STORE.update_cell(
                int(match.group("row_id")),
                field,
                payload.get("value"),
                expected_version,
                name,
                client_id,
            )
        except ValueError as exc:
            self._send_error_json(HTTPStatus.BAD_REQUEST, str(exc))
            return
        except KeyError as exc:
            self._send_error_json(HTTPStatus.NOT_FOUND, str(exc.args[0]))
            return
        if result["conflict"]:
            self._send_json(HTTPStatus.CONFLICT, result)
            return
        touch_presence(client_id, name)
        broadcast("cell", result)
        broadcast("presence", {"presence": presence_snapshot()})
        self._send_json(HTTPStatus.OK, result)

    def do_POST(self) -> None:  # noqa: N802
        path = urlparse(self.path).path
        if path.startswith("/api/v2/"):
            self._v2_post(path)
            return
        try:
            payload = self._json_body()
            client_id, name = normalize_identity(payload.get("client_id"), payload.get("name"))
        except ValueError as exc:
            self._send_error_json(HTTPStatus.BAD_REQUEST, str(exc))
            return

        if path == "/api/presence":
            touch_presence(client_id, name, payload.get("cell"))
            present = presence_snapshot()
            broadcast("presence", {"presence": present})
            self._send_json(HTTPStatus.OK, {"presence": present})
            return
        if path == "/api/cells/batch":
            updates = payload.get("updates")
            if not isinstance(updates, list):
                self._send_error_json(HTTPStatus.BAD_REQUEST, "粘贴内容格式不正确")
                return
            try:
                events = STORE.batch_update_cells(updates, name, client_id)
            except ValueError as exc:
                self._send_error_json(HTTPStatus.BAD_REQUEST, str(exc))
                return
            if events and events[0].get("conflict"):
                self._send_json(HTTPStatus.CONFLICT, events[0])
                return
            touch_presence(client_id, name)
            for event in events:
                broadcast("cell", event)
            broadcast("presence", {"presence": presence_snapshot()})
            self._send_json(HTTPStatus.OK, {"updated": len(events), "events": events})
            return
        if path == "/api/ops/batch":
            operations = payload.get("operations")
            if not isinstance(operations, list) or len(operations) > 100_000:
                self._send_error_json(HTTPStatus.BAD_REQUEST, "操作批次格式不正确")
                return
            try:
                base_revision = int(payload.get("base_revision"))
                events = STORE.apply_operations(operations, name, client_id, base_revision)
            except (TypeError, ValueError, KeyError) as exc:
                self._send_error_json(HTTPStatus.BAD_REQUEST, str(exc))
                return
            conflict = next((event for event in events if event.get("conflict")), None)
            if conflict:
                self._send_json(HTTPStatus.CONFLICT, {
                    "error": "Cell version conflict",
                    "revision": STORE.state["revision"],
                    "event": conflict,
                })
                return
            touch_presence(client_id, name)
            for event in events:
                if "activity" in event:
                    broadcast("cell", event)
                elif "layout" in event:
                    broadcast("layout", event)
            broadcast("presence", {"presence": presence_snapshot()})
            self._send_json(HTTPStatus.OK, {"revision": STORE.state["revision"], "updated": len(events), "events": events})
            return
        if path == "/api/rows":
            after_row_id = payload.get("after_row_id")
            result = STORE.add_row(
                name,
                client_id,
                int(after_row_id) if after_row_id is not None else None,
            )
            touch_presence(client_id, name)
            broadcast("row", result)
            broadcast("presence", {"presence": presence_snapshot()})
            self._send_json(HTTPStatus.CREATED, result)
            return
        if path == "/api/grid/ensure":
            try:
                count = int(payload.get("count", 100))
                result = STORE.ensure_rows(count, name, client_id)
            except (TypeError, ValueError) as exc:
                self._send_error_json(HTTPStatus.BAD_REQUEST, str(exc))
                return
            touch_presence(client_id, name)
            broadcast("grid", result)
            self._send_json(HTTPStatus.OK, result)
            return
        if path == "/api/layout":
            try:
                result = STORE.update_layout(
                    str(payload.get("action") or ""),
                    payload,
                    name,
                    client_id,
                )
            except (TypeError, ValueError) as exc:
                self._send_error_json(HTTPStatus.BAD_REQUEST, str(exc))
                return
            touch_presence(client_id, name)
            broadcast("layout", result)
            self._send_json(HTTPStatus.OK, result)
            return
        self._send_error_json(HTTPStatus.NOT_FOUND, "接口不存在")

    def do_DELETE(self) -> None:  # noqa: N802
        path = urlparse(self.path).path
        if path.startswith("/api/v2/"):
            self._v2_delete(path)
            return
        if path != "/api/rows":
            self._send_error_json(HTTPStatus.NOT_FOUND, "接口不存在")
            return
        try:
            payload = self._json_body()
            client_id, name = normalize_identity(payload.get("client_id"), payload.get("name"))
            row_ids = payload.get("row_ids")
            if not isinstance(row_ids, list):
                raise ValueError("请选择要删除的行")
            result = STORE.delete_rows(row_ids, name, client_id)
        except (TypeError, ValueError) as exc:
            self._send_error_json(HTTPStatus.BAD_REQUEST, str(exc))
            return
        touch_presence(client_id, name)
        broadcast("rows_deleted", result)
        self._send_json(HTTPStatus.OK, result)


def run() -> None:
    server = ThreadingHTTPServer((HOST, PORT), GridBookHandler)
    server.daemon_threads = True
    print(f"GridBook service listening on http://{HOST}:{PORT}", flush=True)
    try:
        server.serve_forever(poll_interval=0.25)
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    run()
