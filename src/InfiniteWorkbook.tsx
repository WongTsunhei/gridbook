import {
  type CSSProperties,
  type ClipboardEvent,
  type KeyboardEvent,
  type MouseEvent as ReactMouseEvent,
  memo,
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import {
  AlignCenter,
  AlignCenterVertical,
  AlignEndVertical,
  AlignJustify,
  AlignLeft,
  AlignRight,
  AlignStartVertical,
  ClipboardPaste,
  Copy,
  Eraser,
  Grid3X3,
  Minus,
  PaintBucket,
  Paintbrush,
  Palette,
  Plus,
  Redo2,
  Scissors,
  Search,
  Snowflake,
  TableCellsMerge,
  TableCellsSplit,
  TableProperties,
  Undo2,
  WrapText,
  X,
} from "lucide-react";
import type { CellBorders, CellFormat, CollaborationIdentity } from "./types";
import { parseHtmlSpreadsheetClipboard, parseSpreadsheetClipboard, serializeSpreadsheetClipboard } from "./clipboard";
import { translateFormula } from "./fill";
import {
  applyV2Operations,
  createV2Sheet,
  deleteV2Sheet,
  getV2Sheets,
  getV2Window,
  renameV2Sheet,
  updateV2Presence,
  v2EventsUrl,
  type V2CellMeta,
  type V2Delta,
  type V2Merge,
  type V2SheetSummary,
  type V2Snapshot,
  type V2UsedRange,
} from "./workbookV2Api";

const CLIENT_KEY = "gridbook-client-id";
const NAME_KEY = "gridbook-display-name";
const ACTIVE_SHEET_KEY = "gridbook-v2-active-sheet";
const DEFAULT_ROW_HEIGHT = 29;
const DEFAULT_COLUMN_WIDTH = 120;
const HEADER_HEIGHT = 32;
const ROW_HEADER_WIDTH = 46;
const MAX_BATCH_CELLS = 100_000;
// At the planned 5,000 × 100 acceptance size the complete rails stay mounted.
// Larger persisted extents use overlapping rail chunks so an unbounded sheet
// cannot turn its two labels-only tracks into hundreds of thousands of DOM
// buttons. The cell body virtualization is intentionally unrelated.
// Navigation keeps a small viewport buffer beyond the used range (about 14
// rows and 18 columns at the default viewport), so the limits include that
// buffer while still covering the requested 5,000 × 100 sheet.
const FULL_RAIL_ROW_LIMIT = 5_200;
const FULL_RAIL_COLUMN_LIMIT = 120;
const RAIL_ROW_CHUNK_SIZE = 2_048;
const RAIL_COLUMN_CHUNK_SIZE = 64;
// Chromium scroll surfaces become unreliable well before their theoretical
// maximum.  Normal workbooks use their full native extent; only very large
// logical sheets are rebased into this physical segment size.
const MAX_NATIVE_SCROLL_PIXELS = 12_000_000;
const FORMULA_NAMES = ["SUM", "AVERAGE", "COUNT", "COUNTA", "MAX", "MIN", "IF", "ROUND"];
const FILL_COLORS = ["#ffffff", "#fff2cc", "#ddebf7", "#e2f0d9", "#fce4d6", "#eadcf8"];
const FONT_COLORS = ["#1f2329", "#d9363e", "#e88000", "#2f9e44", "#2468f2", "#7a3bd2"];

type CellPosition = { row: number; col: number };
type Bounds = { top: number; bottom: number; left: number; right: number };
type Selection = { anchor: CellPosition; focus: CellPosition };
type EditingCell = { row: number; col: number; value: string } | null;
type NavigationExtent = { rows: number; cols: number };
type AxisSegment = { start: number; end: number; pixelStart: number; pixelEnd: number };

// ScrollTimeline is implemented by the current Chromium engines used for
// acceptance, but it is not yet declared by the TypeScript DOM library used
// by this project. Keep the declaration local to the two frozen rails.
interface NativeScrollTimelineOptions {
  source: Element;
  axis: "x" | "y";
}

interface NativeScrollTimelineConstructor {
  new(options: NativeScrollTimelineOptions): AnimationTimeline;
}

type ScrollTimelineWindow = Window & { ScrollTimeline?: NativeScrollTimelineConstructor };

interface SparseSheet {
  id: string;
  name: string;
  usedRange: V2UsedRange;
  contentRange: V2UsedRange;
  cells: Record<string, string>;
  meta: Record<string, V2CellMeta>;
  formats: Record<string, CellFormat>;
  rowHeights: Record<string, number>;
  columnWidths: Record<string, number>;
  merges: V2Merge[];
}

interface Presence {
  client_id: string;
  name: string;
  sheet_id: string;
  cell: CellPosition | null;
}

interface HistoryEntry {
  undo: Array<Record<string, unknown>>;
  redo: Array<Record<string, unknown>>;
}

function cellKey(row: number, col: number): string { return `${row}:${col}`; }
function parseCoordinate(key: string): CellPosition {
  const [row, col] = key.split(":").map(Number);
  return { row, col };
}

function boundsOf(selection: Selection): Bounds {
  return {
    top: Math.min(selection.anchor.row, selection.focus.row),
    bottom: Math.max(selection.anchor.row, selection.focus.row),
    left: Math.min(selection.anchor.col, selection.focus.col),
    right: Math.max(selection.anchor.col, selection.focus.col),
  };
}

function clamp(value: number, minimum: number, maximum: number): number {
  return Math.max(minimum, Math.min(maximum, value));
}

function expandedWindow(window: V2UsedRange): V2UsedRange {
  const rows = Math.max(1, window.bottom - window.top + 1);
  const cols = Math.max(1, window.right - window.left + 1);
  return {
    top: Math.max(0, window.top - rows * 2),
    left: Math.max(0, window.left - cols * 2),
    bottom: window.bottom + rows * 2,
    right: window.right + cols * 2,
  };
}

function keyInBounds(key: string, bounds: V2UsedRange): boolean {
  const [row, col] = key.split(":").map(Number);
  return row >= bounds.top && row <= bounds.bottom && col >= bounds.left && col <= bounds.right;
}

function pruneWindowCache(sheet: SparseSheet, window: V2UsedRange): SparseSheet {
  const retain = expandedWindow(window);
  const trim = <T,>(source: Record<string, T>) => Object.fromEntries(Object.entries(source).filter(([key]) => keyInBounds(key, retain))) as Record<string, T>;
  return {
    ...sheet,
    cells: trim(sheet.cells),
    meta: trim(sheet.meta),
    formats: trim(sheet.formats),
    // Axis overrides are sparse global metrics.  Keeping them lets offsets
    // stay exact even after the cell cache has moved far beyond a resized row
    // or column.
    rowHeights: sheet.rowHeights,
    columnWidths: sheet.columnWidths,
    merges: sheet.merges.filter((merge) => !(merge.end.row < retain.top || merge.start.row > retain.bottom || merge.end.col < retain.left || merge.start.col > retain.right)),
  };
}

function columnName(index: number): string {
  let value = index + 1;
  let name = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}

function columnIndex(name: string): number {
  let value = 0;
  for (const char of name.toUpperCase()) value = value * 26 + char.charCodeAt(0) - 64;
  return value - 1;
}

function createClientId(): string {
  const current = localStorage.getItem(CLIENT_KEY);
  if (current) return current;
  const suffix = typeof crypto.randomUUID === "function" ? crypto.randomUUID() : `${Date.now()}-${Math.random().toString(16).slice(2)}`;
  const next = `client-${suffix}`;
  localStorage.setItem(CLIENT_KEY, next);
  return next;
}

function blankSheet(id: string, name: string): SparseSheet {
  return { id, name, usedRange: { top: 0, left: 0, bottom: -1, right: -1 }, contentRange: { top: 0, left: 0, bottom: -1, right: -1 }, cells: {}, meta: {}, formats: {}, rowHeights: {}, columnWidths: {}, merges: [] };
}

function isMeaningfulFormat(format: CellFormat | undefined): boolean {
  if (!format) return false;
  return Object.values(format).some((value) => value !== undefined && value !== false && value !== "" && value !== "none" && value !== "default");
}

function cloneSheetWindow(snapshot: V2Snapshot): SparseSheet {
  const data = snapshot.sheet;
  return {
    id: data.id,
    name: data.name,
    usedRange: data.used_range,
    contentRange: data.content_range,
    cells: { ...data.cells },
    meta: { ...data.meta },
    formats: { ...data.formats },
    rowHeights: { ...data.row_heights },
    columnWidths: { ...data.column_widths },
    merges: [...data.merges],
  };
}

function keyInWindow(key: string, window: V2UsedRange): boolean {
  const [row, col] = key.split(":").map(Number);
  return row >= window.top && row <= window.bottom && col >= window.left && col <= window.right;
}

function windowContains(container: V2UsedRange, target: V2UsedRange): boolean {
  return target.top >= container.top
    && target.bottom <= container.bottom
    && target.left >= container.left
    && target.right <= container.right;
}

function mergeWindow(previous: SparseSheet | null, snapshot: V2Snapshot): SparseSheet {
  if (!previous || previous.id !== snapshot.sheet.id) return cloneSheetWindow(snapshot);
  const replace = <T,>(source: Record<string, T>, incoming: Record<string, T>) => {
    const next = { ...source };
    Object.keys(next).forEach((key) => { if (keyInWindow(key, snapshot.window)) delete next[key]; });
    return { ...next, ...incoming };
  };
  const nextMerges = [
    ...previous.merges.filter((merge) => merge.end.row < snapshot.window.top || merge.start.row > snapshot.window.bottom || merge.end.col < snapshot.window.left || merge.start.col > snapshot.window.right),
    ...snapshot.sheet.merges,
  ];
  return pruneWindowCache({
    ...previous,
    name: snapshot.sheet.name,
    usedRange: snapshot.sheet.used_range,
    contentRange: snapshot.sheet.content_range,
    cells: replace(previous.cells, snapshot.sheet.cells),
    meta: replace(previous.meta, snapshot.sheet.meta),
    formats: replace(previous.formats, snapshot.sheet.formats),
    rowHeights: { ...snapshot.sheet.row_heights },
    columnWidths: { ...snapshot.sheet.column_widths },
    merges: Array.from(new Map(nextMerges.map((merge) => [merge.id, merge])).values()),
  }, snapshot.window);
}

class SparseAxisMetrics {
  private readonly keys: number[];
  private readonly prefix: number[];
  private readonly deltas: number[];

  constructor(private readonly fallback: number, overrides: Record<string, number>, private readonly scale: number) {
    this.keys = Object.keys(overrides).map(Number).filter((index) => Number.isFinite(index) && index >= 0).sort((a, b) => a - b);
    this.deltas = this.keys.map((index) => ((overrides[String(index)] ?? fallback) - fallback) * scale);
    this.prefix = [];
    let total = 0;
    this.deltas.forEach((delta) => { total += delta; this.prefix.push(total); });
  }

  size(index: number): number {
    const position = this.find(index);
    return (position >= 0 && this.keys[position] === index ? this.fallback + this.deltas[position] / this.scale : this.fallback) * this.scale;
  }

  offset(index: number): number {
    if (index <= 0) return 0;
    const position = this.upperBound(index - 1);
    return index * this.fallback * this.scale + (position >= 0 ? this.prefix[position] : 0);
  }

  span(start: number, endInclusive: number): number { return this.offset(endInclusive + 1) - this.offset(start); }

  indexAt(offset: number, count: number): number {
    let low = 0;
    let high = Math.max(0, count - 1);
    while (low < high) {
      const middle = Math.ceil((low + high) / 2);
      if (this.offset(middle) <= offset) low = middle;
      else high = middle - 1;
    }
    return Math.max(0, Math.min(count - 1, low));
  }

  private find(value: number): number {
    let low = 0; let high = this.keys.length - 1;
    while (low <= high) { const middle = (low + high) >> 1; if (this.keys[middle] === value) return middle; if (this.keys[middle] < value) low = middle + 1; else high = middle - 1; }
    return -1;
  }

  private upperBound(value: number): number {
    let low = 0; let high = this.keys.length - 1; let result = -1;
    while (low <= high) { const middle = (low + high) >> 1; if (this.keys[middle] <= value) { result = middle; low = middle + 1; } else high = middle - 1; }
    return result;
  }
}

function segmentFor(metrics: SparseAxisMetrics, count: number, requestedStart: number): AxisSegment {
  const totalPixels = metrics.offset(count);
  if (totalPixels <= MAX_NATIVE_SCROLL_PIXELS) {
    return { start: 0, end: count, pixelStart: 0, pixelEnd: totalPixels };
  }
  const start = clamp(requestedStart, 0, Math.max(0, count - 1));
  const pixelStart = metrics.offset(start);
  const end = Math.min(count, metrics.indexAt(Math.min(totalPixels - 1, pixelStart + MAX_NATIVE_SCROLL_PIXELS), count) + 1);
  return { start, end, pixelStart, pixelEnd: metrics.offset(end) };
}

function normalizeWheelDelta(event: WheelEvent, viewportSize: number): { row: number; col: number } {
  const multiplier = event.deltaMode === WheelEvent.DOM_DELTA_LINE ? 16 : event.deltaMode === WheelEvent.DOM_DELTA_PAGE ? viewportSize : 1;
  const row = event.deltaY * multiplier;
  const rawCol = event.deltaX * multiplier;
  const col = event.shiftKey && Math.abs(rawCol) < Math.abs(row) ? row : rawCol;
  return { row: event.shiftKey ? 0 : row, col };
}

function rangeFromFormula(value: string): Array<{ top: number; bottom: number; left: number; right: number; color: string }> {
  const colors = ["#3d73f1", "#d86a16", "#0d9c78", "#8a55c7"];
  const ranges: Array<{ top: number; bottom: number; left: number; right: number; color: string }> = [];
  const pattern = /\$?([A-Z]{1,8})\$?(\d+)(?::\$?([A-Z]{1,8})\$?(\d+))?/g;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(value))) {
    const left = columnIndex(match[1]); const top = Number(match[2]) - 1;
    const right = match[3] ? columnIndex(match[3]) : left; const bottom = match[4] ? Number(match[4]) - 1 : top;
    if (left >= 0 && top >= 0 && right >= 0 && bottom >= 0) ranges.push({ top: Math.min(top, bottom), bottom: Math.max(top, bottom), left: Math.min(left, right), right: Math.max(left, right), color: colors[ranges.length % colors.length] });
  }
  return ranges;
}

function simpleFormulaDisplay(raw: string, read: (row: number, col: number) => string): string {
  if (!raw.trimStart().startsWith("=")) return raw;
  const functionMatch = raw.trim().match(/^=\s*(SUM|AVERAGE|COUNT|COUNTA|MAX|MIN)\s*\(\s*([A-Z]+\d+)\s*:\s*([A-Z]+\d+)\s*\)$/i);
  if (!functionMatch) return raw;
  const parse = (reference: string) => {
    const match = reference.match(/^([A-Z]+)(\d+)$/i);
    return match ? { col: columnIndex(match[1]), row: Number(match[2]) - 1 } : null;
  };
  const start = parse(functionMatch[2]); const end = parse(functionMatch[3]);
  if (!start || !end) return raw;
  const values: string[] = [];
  for (let row = Math.min(start.row, end.row); row <= Math.max(start.row, end.row); row += 1) for (let col = Math.min(start.col, end.col); col <= Math.max(start.col, end.col); col += 1) values.push(read(row, col));
  const numbers = values.map((value) => Number(value.replaceAll(",", ""))).filter((value) => Number.isFinite(value));
  switch (functionMatch[1].toUpperCase()) {
    case "SUM": return String(numbers.reduce((sum, value) => sum + value, 0));
    case "AVERAGE": return String(numbers.length ? numbers.reduce((sum, value) => sum + value, 0) / numbers.length : 0);
    case "COUNT": return String(numbers.length);
    case "COUNTA": return String(values.filter(Boolean).length);
    case "MAX": return String(numbers.length ? Math.max(...numbers) : 0);
    case "MIN": return String(numbers.length ? Math.min(...numbers) : 0);
    default: return raw;
  }
}

function formatCellValue(value: string, format: CellFormat | undefined, read: (row: number, col: number) => string): string {
  const raw = simpleFormulaDisplay(value, read);
  if (!raw || raw.trimStart().startsWith("=")) return raw;
  if (format?.number_format === "percentage") {
    const number = Number(raw.replace("%", ""));
    return Number.isFinite(number) ? `${number * (raw.endsWith("%") ? 1 : 100)}%` : raw;
  }
  if (format?.number_format === "number") {
    const number = Number(raw.replaceAll(",", ""));
    return Number.isFinite(number) ? new Intl.NumberFormat("zh-CN", { maximumFractionDigits: 2 }).format(number) : raw;
  }
  return raw;
}

function getCellBorders(format: CellFormat | undefined): CSSProperties {
  const edge = (format?.border && !format?.borders ? { style: "solid", width: 1, color: "#66717d" } : undefined) as CellBorders["top"];
  const styleFor = (value: CellBorders["top"] | undefined) => value ? `${value.width}px ${value.style} ${value.color}` : undefined;
  return {
    borderTop: styleFor(format?.borders?.top ?? edge),
    borderRight: styleFor(format?.borders?.right ?? edge),
    borderBottom: styleFor(format?.borders?.bottom ?? edge),
    borderLeft: styleFor(format?.borders?.left ?? edge),
  };
}

interface InfiniteGridProps {
  sheet: SparseSheet;
  navigation: { rows: number; cols: number };
  selection: Selection;
  selections: Selection[];
  editing: EditingCell;
  identity: CollaborationIdentity;
  presence: Presence[];
  zoom: number;
  formulaDraft: string;
  onViewport: (rows: number, cols: number) => void;
  onVisibleRange: (range: V2UsedRange) => void;
  onFetchWindow: (top: number, left: number, rows: number, cols: number) => void;
  onWheelExtend: (axis: "row" | "col", restore: number) => void;
  onFillExtend: (axis: "row" | "col") => void;
  onSelect: (row: number, col: number, extend: boolean, additive: boolean, grow?: boolean) => void;
  onContext: (event: ReactMouseEvent, row: number, col: number) => void;
  onDragSelect: (row: number, col: number) => void;
  onSelectRow: (row: number, extend: boolean, additive: boolean) => void;
  onSelectColumn: (col: number, extend: boolean, additive: boolean) => void;
  onSelectAll: () => void;
  onBeginEdit: (row: number, col: number, seed?: string) => void;
  onEditChange: (value: string) => void;
  onCommitEdit: () => void;
  onCancelEdit: () => void;
  onResizeRow: (row: number, height: number, persist: boolean) => void;
  onResizeColumn: (col: number, width: number, persist: boolean) => void;
  onFill: (source: Selection, target: CellPosition, copy: boolean) => void;
  onFormulaReference: (row: number, col: number, extend: boolean) => boolean;
}

interface InfiniteGridColumnRailItemProps {
  col: number;
  left: number;
  width: number;
  height: number;
  selected: boolean;
  draggingRef: { current: boolean };
  onSelectColumn: (col: number, extend: boolean, additive: boolean) => void;
  onResize: (col: number, event: ReactMouseEvent) => void;
}

const InfiniteGridColumnRailItem = memo(function InfiniteGridColumnRailItem({
  col, left, width, height, selected, draggingRef, onSelectColumn, onResize,
}: InfiniteGridColumnRailItemProps) {
  return <button
    className={selected ? "infinite-grid-col-head selected" : "infinite-grid-col-head"}
    style={{ left, width, height }}
    onMouseDown={(event) => {
      event.preventDefault();
      draggingRef.current = true;
      onSelectColumn(col, event.shiftKey, event.ctrlKey || event.metaKey);
    }}
    onMouseEnter={() => { if (draggingRef.current) onSelectColumn(col, true, false); }}
    type="button"
  ><span>{columnName(col)}</span><i onMouseDown={(event) => onResize(col, event)} /></button>;
});

interface InfiniteGridRowRailItemProps {
  row: number;
  top: number;
  width: number;
  height: number;
  selected: boolean;
  draggingRef: { current: boolean };
  onSelectRow: (row: number, extend: boolean, additive: boolean) => void;
  onResize: (row: number, event: ReactMouseEvent) => void;
}

const InfiniteGridRowRailItem = memo(function InfiniteGridRowRailItem({
  row, top, width, height, selected, draggingRef, onSelectRow, onResize,
}: InfiniteGridRowRailItemProps) {
  return <button
    className={selected ? "infinite-grid-row-head selected" : "infinite-grid-row-head"}
    style={{ top, width, height }}
    onMouseDown={(event) => {
      event.preventDefault();
      draggingRef.current = true;
      onSelectRow(row, event.shiftKey, event.ctrlKey || event.metaKey);
    }}
    onMouseEnter={() => { if (draggingRef.current) onSelectRow(row, true, false); }}
    type="button"
  >{row + 1}<i onMouseDown={(event) => onResize(row, event)} /></button>;
});

function InfiniteGrid({
  sheet, navigation, selection, selections, editing, identity, presence, zoom, formulaDraft,
  onViewport, onVisibleRange, onFetchWindow, onWheelExtend, onFillExtend, onSelect, onContext, onDragSelect, onSelectRow, onSelectColumn, onSelectAll,
  onBeginEdit, onEditChange, onCommitEdit, onCancelEdit, onResizeRow, onResizeColumn, onFill, onFormulaReference,
}: InfiniteGridProps) {
  const shellRef = useRef<HTMLDivElement>(null);
  const scrollRef = useRef<HTMLDivElement>(null);
  const headerTrackRef = useRef<HTMLDivElement>(null);
  const rowTrackRef = useRef<HTMLDivElement>(null);
  const headerRailAnimationRef = useRef<Animation | null>(null);
  const rowRailAnimationRef = useRef<Animation | null>(null);
  const scrollFrameRef = useRef<number | null>(null);
  const scrollPositionRef = useRef({ top: 0, left: 0 });
  const visibleRangeRef = useRef<V2UsedRange>({ top: -1, left: -1, bottom: -1, right: -1 });
  const pendingRestoreRef = useRef<{ row?: number; col?: number } | null>(null);
  const fillExtendLockRef = useRef<{ row: number; col: number }>({ row: -1, col: -1 });
  const fillTargetRef = useRef<CellPosition | null>(null);
  const draggingRef = useRef(false);
  const [scroll, setScroll] = useState({ top: 0, left: 0 });
  const [viewport, setViewport] = useState({ width: 1100, height: 520 });
  const [fillTarget, setFillTarget] = useState<CellPosition | null>(null);
  const [segmentStart, setSegmentStart] = useState({ row: 0, col: 0 });
  const fillDragRef = useRef<{ source: Selection; copy: boolean } | null>(null);
  const scrollTimelineConstructor = useMemo<NativeScrollTimelineConstructor | null>(() => {
    if (typeof window === "undefined") return null;
    const candidate = (window as ScrollTimelineWindow).ScrollTimeline;
    return typeof candidate === "function" ? candidate : null;
  }, []);
  const usesNativeRailTimelines = scrollTimelineConstructor !== null;
  const scale = Math.max(0.5, Math.min(2, zoom / 100));
  const rows = useMemo(() => new SparseAxisMetrics(DEFAULT_ROW_HEIGHT, sheet.rowHeights, scale), [scale, sheet.rowHeights]);
  const cols = useMemo(() => new SparseAxisMetrics(DEFAULT_COLUMN_WIDTH, sheet.columnWidths, scale), [scale, sheet.columnWidths]);
  const rowSegment = useMemo(() => segmentFor(rows, navigation.rows, segmentStart.row), [navigation.rows, rows, segmentStart.row]);
  const colSegment = useMemo(() => segmentFor(cols, navigation.cols, segmentStart.col), [cols, navigation.cols, segmentStart.col]);
  const bodyWidth = colSegment.pixelEnd - colSegment.pixelStart;
  const bodyHeight = rowSegment.pixelEnd - rowSegment.pixelStart;
  const rowOverscan = Math.max(10, Math.ceil(viewport.height / Math.max(rows.size(0), 1)) + 6);
  const colOverscan = Math.max(6, Math.ceil(viewport.width / Math.max(cols.size(0), 1)) + 4);
  const firstRow = Math.max(rowSegment.start, rows.indexAt(rowSegment.pixelStart + scroll.top, navigation.rows) - rowOverscan);
  const lastRow = Math.min(rowSegment.end - 1, rows.indexAt(rowSegment.pixelStart + scroll.top + viewport.height, navigation.rows) + rowOverscan);
  const firstCol = Math.max(colSegment.start, cols.indexAt(colSegment.pixelStart + scroll.left, navigation.cols) - colOverscan);
  const lastCol = Math.min(colSegment.end - 1, cols.indexAt(colSegment.pixelStart + scroll.left + viewport.width, navigation.cols) + colOverscan);
  const visibleRows = useMemo(() => Array.from({ length: Math.max(0, lastRow - firstRow + 1) }, (_, index) => firstRow + index), [firstRow, lastRow]);
  const visibleCols = useMemo(() => Array.from({ length: Math.max(0, lastCol - firstCol + 1) }, (_, index) => firstCol + index), [firstCol, lastCol]);
  // The rail labels intentionally have their own stable render set. The
  // virtual cell body remains unchanged, while these lightweight buttons are
  // ready before a scroll-frame React update asks the body for a new window.
  // For the genuinely unbounded default extent, retain one previous and one
  // next chunk so a chunk swap happens well before its labels reach the
  // viewport. Within the planned 5,000 × 100 range this resolves to a full
  // rail, with no scroll-derived rail render at all.
  const railRowStart = navigation.rows <= FULL_RAIL_ROW_LIMIT
    ? 0
    : Math.max(0, (Math.floor(firstRow / RAIL_ROW_CHUNK_SIZE) - 1) * RAIL_ROW_CHUNK_SIZE);
  const railRowEnd = navigation.rows <= FULL_RAIL_ROW_LIMIT
    ? navigation.rows
    : Math.min(navigation.rows, (Math.floor(firstRow / RAIL_ROW_CHUNK_SIZE) + 2) * RAIL_ROW_CHUNK_SIZE);
  const railColumnStart = navigation.cols <= FULL_RAIL_COLUMN_LIMIT
    ? 0
    : Math.max(0, (Math.floor(firstCol / RAIL_COLUMN_CHUNK_SIZE) - 1) * RAIL_COLUMN_CHUNK_SIZE);
  const railColumnEnd = navigation.cols <= FULL_RAIL_COLUMN_LIMIT
    ? navigation.cols
    : Math.min(navigation.cols, (Math.floor(firstCol / RAIL_COLUMN_CHUNK_SIZE) + 2) * RAIL_COLUMN_CHUNK_SIZE);
  const railRows = useMemo(() => Array.from({ length: Math.max(0, railRowEnd - railRowStart) }, (_, index) => railRowStart + index), [railRowEnd, railRowStart]);
  const railCols = useMemo(() => Array.from({ length: Math.max(0, railColumnEnd - railColumnStart) }, (_, index) => railColumnStart + index), [railColumnEnd, railColumnStart]);
  const allSelections = useMemo(() => [...selections, selection], [selection, selections]);
  const formulaRanges = useMemo(() => rangeFromFormula(formulaDraft), [formulaDraft]);

  const mergeGeometry = useMemo(() => {
    const anchors = new Map<string, V2Merge>();
    const hidden = new Set<string>();
    sheet.merges.forEach((merge) => {
      anchors.set(cellKey(merge.start.row, merge.start.col), merge);
      for (let row = merge.start.row; row <= merge.end.row; row += 1) for (let col = merge.start.col; col <= merge.end.col; col += 1) {
        if (row !== merge.start.row || col !== merge.start.col) hidden.add(cellKey(row, col));
      }
    });
    return { anchors, hidden };
  }, [sheet.merges]);

  const cancelRailScrollTimelines = useCallback(() => {
    headerRailAnimationRef.current?.cancel();
    rowRailAnimationRef.current?.cancel();
    headerRailAnimationRef.current = null;
    rowRailAnimationRef.current = null;
  }, []);

  // This is the only fallback path. It is deliberately isolated from the
  // normal virtual-body scroll listener below: Chromium never writes rail
  // transforms from JavaScript while it is scrolling.
  const syncFallbackRailTransforms = useCallback((target: HTMLDivElement) => {
    if (headerTrackRef.current) headerTrackRef.current.style.transform = `translate3d(${-target.scrollLeft}px,0,0)`;
    if (rowTrackRef.current) rowTrackRef.current.style.transform = `translate3d(0,${-target.scrollTop}px,0)`;
  }, []);

  const rebuildRailScrollTimelines = useCallback(() => {
    cancelRailScrollTimelines();
    const target = scrollRef.current;
    const headerTrack = headerTrackRef.current;
    const rowTrack = rowTrackRef.current;
    if (!target || !headerTrack || !rowTrack || !scrollTimelineConstructor) return;

    // The scroll range, not a JavaScript callback, is the time source. The
    // computed translation at progress p is exactly -p * maxScroll.
    const maxScrollX = Math.max(0, target.scrollWidth - target.clientWidth);
    const maxScrollY = Math.max(0, target.scrollHeight - target.clientHeight);
    const horizontalTimeline = new scrollTimelineConstructor({ source: target, axis: "x" });
    const verticalTimeline = new scrollTimelineConstructor({ source: target, axis: "y" });

    // Clear a transform left by a non-Chromium fallback before the native
    // animation takes ownership. This executes only during a layout rebuild,
    // never as part of scrolling.
    headerTrack.style.transform = "";
    rowTrack.style.transform = "";
    headerRailAnimationRef.current = headerTrack.animate([
      { transform: "translate3d(0px, 0px, 0px)" },
      { transform: `translate3d(${-maxScrollX}px, 0px, 0px)` },
    ], {
      duration: 1,
      easing: "linear",
      fill: "both",
      composite: "replace",
      timeline: horizontalTimeline,
    });
    rowRailAnimationRef.current = rowTrack.animate([
      { transform: "translate3d(0px, 0px, 0px)" },
      { transform: `translate3d(0px, ${-maxScrollY}px, 0px)` },
    ], {
      duration: 1,
      easing: "linear",
      fill: "both",
      composite: "replace",
      timeline: verticalTimeline,
    });
  }, [cancelRailScrollTimelines, scrollTimelineConstructor]);

  useEffect(() => {
    setSegmentStart((current) => {
      const next = {
        row: rows.offset(navigation.rows) <= MAX_NATIVE_SCROLL_PIXELS ? 0 : clamp(current.row, 0, Math.max(0, navigation.rows - 1)),
        col: cols.offset(navigation.cols) <= MAX_NATIVE_SCROLL_PIXELS ? 0 : clamp(current.col, 0, Math.max(0, navigation.cols - 1)),
      };
      return next.row === current.row && next.col === current.col ? current : next;
    });
  }, [cols, navigation.cols, navigation.rows, rows]);

  useEffect(() => {
    const next: { row?: number; col?: number } = {};
    let needsRebase = false;
    if (selection.focus.row < rowSegment.start || selection.focus.row >= rowSegment.end) {
      const targetOffset = rows.offset(selection.focus.row);
      next.row = targetOffset;
      const start = rows.indexAt(Math.max(0, targetOffset - MAX_NATIVE_SCROLL_PIXELS * 0.28), navigation.rows);
      setSegmentStart((current) => current.row === start ? current : { ...current, row: start });
      needsRebase = true;
    }
    if (selection.focus.col < colSegment.start || selection.focus.col >= colSegment.end) {
      const targetOffset = cols.offset(selection.focus.col);
      next.col = targetOffset;
      const start = cols.indexAt(Math.max(0, targetOffset - MAX_NATIVE_SCROLL_PIXELS * 0.28), navigation.cols);
      setSegmentStart((current) => current.col === start ? current : { ...current, col: start });
      needsRebase = true;
    }
    if (needsRebase) pendingRestoreRef.current = { ...pendingRestoreRef.current, ...next };
  // A wheel-driven segment rebase must not snap back to an older selection.
  // Selection navigation itself changes the focus coordinates, which is the
  // only signal that should request a follow-selection segment move.
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selection.focus.col, selection.focus.row]);

  useLayoutEffect(() => {
    const target = scrollRef.current;
    const pending = pendingRestoreRef.current;
    if (!target || !pending) return;
    pendingRestoreRef.current = null;
    if (pending.row !== undefined) target.scrollTop = clamp(pending.row - rowSegment.pixelStart, 0, Math.max(0, target.scrollHeight - target.clientHeight));
    if (pending.col !== undefined) target.scrollLeft = clamp(pending.col - colSegment.pixelStart, 0, Math.max(0, target.scrollWidth - target.clientWidth));
    const next = { top: target.scrollTop, left: target.scrollLeft };
    scrollPositionRef.current = next;
    setScroll(next);
  }, [bodyHeight, bodyWidth, colSegment.pixelStart, rowSegment.pixelStart]);

  useLayoutEffect(() => {
    const target = scrollRef.current;
    if (!target) return undefined;
    const update = () => {
      const next = { width: target.clientWidth, height: target.clientHeight };
      setViewport(next);
      onViewport(Math.max(1, Math.ceil(next.height / Math.max(rows.size(0), 1))), Math.max(1, Math.ceil(next.width / Math.max(cols.size(0), 1))));
      rebuildRailScrollTimelines();
    };
    update();
    const observer = new ResizeObserver(update);
    observer.observe(target);
    return () => observer.disconnect();
  }, [cols, onViewport, rebuildRailScrollTimelines, rows]);

  // Geometry changes are the only native-animation rebuild trigger. In
  // particular, scroll, wheel, pointer movement and selection drag are not
  // dependencies of this effect.
  useLayoutEffect(() => {
    rebuildRailScrollTimelines();
  }, [bodyHeight, bodyWidth, colSegment.pixelStart, navigation.cols, navigation.rows, rebuildRailScrollTimelines, rowSegment.pixelStart, scale]);

  useEffect(() => () => cancelRailScrollTimelines(), [cancelRailScrollTimelines]);

  // Non-Chromium browsers retain the older JavaScript path solely as a
  // compatibility fallback. Target browsers never install this listener.
  useLayoutEffect(() => {
    const target = scrollRef.current;
    if (!target || usesNativeRailTimelines) return undefined;
    const sync = () => syncFallbackRailTransforms(target);
    sync();
    target.addEventListener("scroll", sync, { passive: true });
    return () => target.removeEventListener("scroll", sync);
  }, [bodyHeight, bodyWidth, syncFallbackRailTransforms, usesNativeRailTimelines]);

  useEffect(() => {
    const target = scrollRef.current;
    if (!target) return undefined;
    const onScroll = () => {
      if (scrollFrameRef.current !== null) return;
      scrollFrameRef.current = requestAnimationFrame(() => {
        scrollFrameRef.current = null;
        const next = { top: target.scrollTop, left: target.scrollLeft };
        scrollPositionRef.current = next;
        setScroll((current) => {
          const currentRow = rows.indexAt(rowSegment.pixelStart + current.top, navigation.rows);
          const currentCol = cols.indexAt(colSegment.pixelStart + current.left, navigation.cols);
          const nextRow = rows.indexAt(rowSegment.pixelStart + next.top, navigation.rows);
          const nextCol = cols.indexAt(colSegment.pixelStart + next.left, navigation.cols);
          return currentRow === nextRow && currentCol === nextCol ? current : next;
        });
      });
    };
    target.addEventListener("scroll", onScroll, { passive: true });
    onScroll();
    return () => {
      target.removeEventListener("scroll", onScroll);
      // An effect rebind can happen while the first frame is queued (for
      // example when the initial sparse window arrives).  Cancelling that
      // frame without resetting this sentinel permanently suppresses all
      // later virtual-window updates: frozen headers keep moving, but rows
      // and columns no longer re-render.  Always clear it together with the
      // cancellation.
      if (scrollFrameRef.current !== null) {
        cancelAnimationFrame(scrollFrameRef.current);
        scrollFrameRef.current = null;
      }
    };
  }, [cols, colSegment.pixelStart, navigation.cols, navigation.rows, rowSegment.pixelStart, rows]);

  // Native wheel expansion changes the canvas dimensions without necessarily
  // emitting a second scroll event. Re-read the browser position so the
  // virtual row/column window is never left at the pre-extension range.
  useEffect(() => {
    const target = scrollRef.current;
    if (!target) return;
    const next = { top: target.scrollTop, left: target.scrollLeft };
    scrollPositionRef.current = next;
    setScroll(next);
  }, [bodyHeight, bodyWidth, navigation.cols, navigation.rows]);

  useEffect(() => {
    const range = {
      top: rows.indexAt(rowSegment.pixelStart + scroll.top, navigation.rows),
      left: cols.indexAt(colSegment.pixelStart + scroll.left, navigation.cols),
      bottom: rows.indexAt(rowSegment.pixelStart + scroll.top + viewport.height, navigation.rows),
      right: cols.indexAt(colSegment.pixelStart + scroll.left + viewport.width, navigation.cols),
    };
    const current = visibleRangeRef.current;
    if (current.top === range.top && current.left === range.left && current.bottom === range.bottom && current.right === range.right) return;
    visibleRangeRef.current = range;
    onVisibleRange(range);
  }, [cols, colSegment.pixelStart, navigation.cols, navigation.rows, onVisibleRange, rowSegment.pixelStart, rows, scroll.left, scroll.top, viewport.height, viewport.width]);

  useEffect(() => {
    onFetchWindow(firstRow, firstCol, Math.min(160, Math.max(1, lastRow - firstRow + 1)), Math.min(72, Math.max(1, lastCol - firstCol + 1)));
  }, [firstCol, firstRow, lastCol, lastRow, onFetchWindow]);

  useEffect(() => {
    const stop = () => { draggingRef.current = false; };
    window.addEventListener("mouseup", stop);
    return () => window.removeEventListener("mouseup", stop);
  }, []);

  useEffect(() => {
    fillExtendLockRef.current = { row: -1, col: -1 };
  }, [navigation.cols, navigation.rows]);

  useEffect(() => {
    const onMove = (event: PointerEvent) => {
      const drag = fillDragRef.current;
      const target = scrollRef.current;
      if (!drag || !target) return;
      const rect = target.getBoundingClientRect();
      const x = event.clientX - rect.left + target.scrollLeft + colSegment.pixelStart;
      const y = event.clientY - rect.top + target.scrollTop + rowSegment.pixelStart;
      const row = rows.indexAt(Math.max(rowSegment.pixelStart, y), navigation.rows);
      const col = cols.indexAt(Math.max(colSegment.pixelStart, x), navigation.cols);
      const next = { row, col };
      fillTargetRef.current = next;
      setFillTarget(next);
      const atBottom = target.scrollTop >= target.scrollHeight - target.clientHeight - 1;
      const atRight = target.scrollLeft >= target.scrollWidth - target.clientWidth - 1;
      if (event.clientY >= rect.bottom - 22 && atBottom && rowSegment.end >= navigation.rows && fillExtendLockRef.current.row !== navigation.rows) {
        fillExtendLockRef.current.row = navigation.rows;
        onFillExtend("row");
      }
      if (event.clientX >= rect.right - 22 && atRight && colSegment.end >= navigation.cols && fillExtendLockRef.current.col !== navigation.cols) {
        fillExtendLockRef.current.col = navigation.cols;
        onFillExtend("col");
      }
    };
    const onUp = () => {
      const drag = fillDragRef.current;
      if (drag && fillTargetRef.current) onFill(drag.source, fillTargetRef.current, drag.copy);
      fillDragRef.current = null;
      fillTargetRef.current = null;
      setFillTarget(null);
    };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
    window.addEventListener("pointercancel", onUp);
    return () => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      window.removeEventListener("pointercancel", onUp);
    };
  }, [colSegment.pixelStart, colSegment.end, navigation.cols, navigation.rows, onFill, onFillExtend, rowSegment.pixelStart, rowSegment.end, rows, cols]);

  const handleNativeWheel = useCallback((event: WheelEvent) => {
    const target = scrollRef.current;
    if (!target) return;
    const delta = normalizeWheelDelta(event, Math.max(target.clientWidth, target.clientHeight));
    const axis = Math.abs(delta.col) > Math.abs(delta.row) ? "col" : "row";
    const amount = axis === "row" ? delta.row : delta.col;
    if (amount <= 0) return;
    const atBottom = target.scrollTop >= target.scrollHeight - target.clientHeight - 1;
    const atRight = target.scrollLeft >= target.scrollWidth - target.clientWidth - 1;
    const atNativeEnd = axis === "row" ? atBottom : atRight;
    if (!atNativeEnd) return;
    event.preventDefault();
    const metrics = axis === "row" ? rows : cols;
    const segment = axis === "row" ? rowSegment : colSegment;
    const count = axis === "row" ? navigation.rows : navigation.cols;
    const logicalOffset = segment.pixelStart + (axis === "row" ? target.scrollTop : target.scrollLeft) + amount;
    if (segment.end < count) {
      const start = metrics.indexAt(Math.max(0, logicalOffset - MAX_NATIVE_SCROLL_PIXELS * 0.28), count);
      pendingRestoreRef.current = { ...pendingRestoreRef.current, [axis]: logicalOffset };
      setSegmentStart((current) => axis === "row" ? { ...current, row: start } : { ...current, col: start });
      return;
    }
    pendingRestoreRef.current = { ...pendingRestoreRef.current, [axis]: logicalOffset };
    onWheelExtend(axis, logicalOffset);
  }, [colSegment, navigation.cols, navigation.rows, onWheelExtend, rowSegment, rows, cols]);

  useEffect(() => {
    const target = scrollRef.current;
    if (!target) return undefined;
    target.addEventListener("wheel", handleNativeWheel, { passive: false });
    return () => target.removeEventListener("wheel", handleNativeWheel);
  }, [handleNativeWheel]);

  const activeBounds = boundsOf(selection);
  const fillBounds = fillTarget ? {
    top: Math.min(activeBounds.top, fillTarget.row), bottom: Math.max(activeBounds.bottom, fillTarget.row),
    left: Math.min(activeBounds.left, fillTarget.col), right: Math.max(activeBounds.right, fillTarget.col),
  } : null;
  const read = useCallback((row: number, col: number) => sheet.cells[cellKey(row, col)] ?? "", [sheet.cells]);
  const remoteAt = useCallback((row: number, col: number) => presence.find((person) => person.client_id !== identity.clientId && person.cell?.row === row && person.cell.col === col), [identity.clientId, presence]);

  const selected = (row: number, col: number) => allSelections.some((item) => {
    const bounds = boundsOf(item);
    return row >= bounds.top && row <= bounds.bottom && col >= bounds.left && col <= bounds.right;
  });

  const referenceAt = (row: number, col: number) => formulaRanges.find((range) => row >= range.top && row <= range.bottom && col >= range.left && col <= range.right);

  const startColumnResize = useCallback((col: number, event: ReactMouseEvent) => {
    event.preventDefault(); event.stopPropagation();
    const initial = cols.size(col) / scale; const start = event.clientX;
    let finalWidth = initial;
    const move = (next: MouseEvent) => { finalWidth = Math.round(Math.max(40, Math.min(1200, initial + next.clientX - start))); onResizeColumn(col, finalWidth, false); };
    const up = () => { window.removeEventListener("mousemove", move); window.removeEventListener("mouseup", up); onResizeColumn(col, finalWidth, true); };
    window.addEventListener("mousemove", move); window.addEventListener("mouseup", up);
  }, [cols, onResizeColumn, scale]);

  const startRowResize = useCallback((row: number, event: ReactMouseEvent) => {
    event.preventDefault(); event.stopPropagation();
    const initial = rows.size(row) / scale; const start = event.clientY;
    let finalHeight = initial;
    const move = (next: MouseEvent) => { finalHeight = Math.round(Math.max(20, Math.min(600, initial + next.clientY - start))); onResizeRow(row, finalHeight, false); };
    const up = () => { window.removeEventListener("mousemove", move); window.removeEventListener("mouseup", up); onResizeRow(row, finalHeight, true); };
    window.addEventListener("mousemove", move); window.addEventListener("mouseup", up);
  }, [onResizeRow, rows, scale]);

  // These nodes deliberately do not depend on scroll, visibleRows or
  // visibleCols. The scroll listener can therefore keep virtualizing only the
  // cell body while the compositor moves a complete, stable rail in the same
  // frame as the native scroll surface.
  const columnRailNodes = useMemo(() => railCols.map((col) => {
    const selectedColumn = allSelections.some((item) => {
      const bounds = boundsOf(item);
      return col >= bounds.left && col <= bounds.right;
    });
    return <InfiniteGridColumnRailItem
      col={col}
      draggingRef={draggingRef}
      height={HEADER_HEIGHT * scale}
      key={`col-${col}`}
      left={cols.offset(col) - colSegment.pixelStart}
      onResize={startColumnResize}
      onSelectColumn={onSelectColumn}
      selected={selectedColumn}
      width={cols.size(col)}
    />;
  }), [allSelections, colSegment.pixelStart, cols, onSelectColumn, railCols, scale, startColumnResize]);

  const rowRailNodes = useMemo(() => railRows.map((row) => {
    const selectedRow = allSelections.some((item) => {
      const bounds = boundsOf(item);
      return row >= bounds.top && row <= bounds.bottom;
    });
    return <InfiniteGridRowRailItem
      draggingRef={draggingRef}
      height={rows.size(row)}
      key={`row-${row}`}
      onResize={startRowResize}
      onSelectRow={onSelectRow}
      row={row}
      selected={selectedRow}
      top={rows.offset(row) - rowSegment.pixelStart}
      width={ROW_HEADER_WIDTH * scale}
    />;
  }), [allSelections, onSelectRow, railRows, rowSegment.pixelStart, rows, scale, startRowResize]);

  return <div className="infinite-grid-shell" ref={shellRef} style={{ "--sheet-zoom": String(scale) } as CSSProperties}>
    <div className="infinite-grid-fixed-head" style={{ left: ROW_HEADER_WIDTH * scale, height: HEADER_HEIGHT * scale }}>
      <div className="infinite-grid-fixed-head-track" data-rail-driver={usesNativeRailTimelines ? "scroll-timeline" : "javascript-fallback"} ref={headerTrackRef} style={{ width: bodyWidth, height: HEADER_HEIGHT * scale }}>
        {columnRailNodes}
      </div>
    </div>
    <button aria-label="全选表格" className="infinite-grid-corner" style={{ width: ROW_HEADER_WIDTH * scale, height: HEADER_HEIGHT * scale }} onMouseDown={(event) => { event.preventDefault(); onSelectAll(); }} type="button" />
    <div className="infinite-grid-fixed-rows" style={{ top: HEADER_HEIGHT * scale, width: ROW_HEADER_WIDTH * scale }}>
      <div className="infinite-grid-fixed-rows-track" data-rail-driver={usesNativeRailTimelines ? "scroll-timeline" : "javascript-fallback"} ref={rowTrackRef} style={{ height: bodyHeight, width: ROW_HEADER_WIDTH * scale }}>
        {rowRailNodes}
      </div>
    </div>
    <div className="infinite-grid-scroll" ref={scrollRef} style={{ top: HEADER_HEIGHT * scale, left: ROW_HEADER_WIDTH * scale }}>
      <div className="infinite-grid-canvas" style={{ width: bodyWidth, height: bodyHeight }}>
        {visibleRows.map((row) => visibleCols.map((col) => {
          const coordinate = cellKey(row, col);
          if (mergeGeometry.hidden.has(coordinate)) return null;
          const merge = mergeGeometry.anchors.get(coordinate);
          const width = merge ? cols.span(col, merge.end.col) : cols.size(col);
          const height = merge ? rows.span(row, merge.end.row) : rows.size(row);
          const format = sheet.formats[coordinate] ?? {};
          const remote = remoteAt(row, col);
          const formulaReference = referenceAt(row, col);
          const value = read(row, col);
          const isEditing = editing?.row === row && editing.col === col;
          const style: CSSProperties = { left: cols.offset(col) - colSegment.pixelStart, top: rows.offset(row) - rowSegment.pixelStart, width, height, backgroundColor: format.fill && format.fill !== "none" ? format.fill : undefined, color: format.font_color && format.font_color !== "default" ? format.font_color : undefined, fontFamily: format.font_family, fontSize: format.font_size ? `${format.font_size * scale}px` : undefined, textAlign: format.align, justifyContent: format.vertical_align === "bottom" ? "flex-end" : format.vertical_align === "top" ? "flex-start" : "center", ...getCellBorders(format), ...(formulaReference ? { "--formula-reference-color": formulaReference.color } : {}) } as CSSProperties;
          const classes = ["infinite-grid-cell", selected(row, col) ? "selected" : "", isEditing ? "editing" : "", format.bold ? "format-bold" : "", format.italic ? "format-italic" : "", format.underline ? "format-underline" : "", format.wrap ? "format-wrap" : "", formulaReference ? "formula-reference-cell" : "", remote ? "remote-cell" : ""].filter(Boolean).join(" ");
          return <div className={classes} data-cell={`${row}:${col}`} key={coordinate} onContextMenu={(event) => onContext(event, row, col)} onMouseDown={(event) => { if (onFormulaReference(row, col, event.shiftKey)) return; draggingRef.current = true; onSelect(row, col, event.shiftKey, event.ctrlKey || event.metaKey); }} onMouseEnter={() => { if (draggingRef.current) onDragSelect(row, col); }} onDoubleClick={() => onBeginEdit(row, col)} style={style}>
            {isEditing ? <input autoFocus className="infinite-grid-editor" onBlur={onCommitEdit} onChange={(event) => onEditChange(event.target.value)} onKeyDown={(event) => { if (event.key === "Enter" || event.key === "Tab") { event.preventDefault(); onCommitEdit(); onSelect(row + (event.key === "Enter" ? 1 : 0), col + (event.key === "Tab" ? (event.shiftKey ? -1 : 1) : 0), false, false, true); } if (event.key === "Escape") { event.preventDefault(); onCancelEdit(); } }} value={editing.value} /> : <button className="infinite-grid-cell-button" onKeyDown={(event) => { if (event.key === "Enter" || event.key === "F2") { event.preventDefault(); onBeginEdit(row, col); } }} type="button">{formatCellValue(value, format, read)}</button>}
            {remote ? <span className="infinite-remote-tag">{remote.name} 正在编辑</span> : null}
          </div>;
        }))}
        <div className="infinite-selection-layer" aria-hidden="true" style={{ width: bodyWidth, height: bodyHeight }}>
          {allSelections.map((item, index) => {
            const bounds = boundsOf(item);
            return <div className={`infinite-selection${bounds.top === bounds.bottom && bounds.left === bounds.right ? " single" : ""}`} key={`selection-${index}`} style={{ left: cols.offset(bounds.left) - colSegment.pixelStart, top: rows.offset(bounds.top) - rowSegment.pixelStart, width: cols.span(bounds.left, bounds.right), height: rows.span(bounds.top, bounds.bottom) }} />;
          })}
          {formulaRanges.map((range, index) => <div className="infinite-formula-range" key={`formula-${index}`} style={{ left: cols.offset(range.left) - colSegment.pixelStart, top: rows.offset(range.top) - rowSegment.pixelStart, width: cols.span(range.left, range.right), height: rows.span(range.top, range.bottom), borderColor: range.color }} />)}
          {fillBounds ? <div className="infinite-fill-preview" style={{ left: cols.offset(fillBounds.left) - colSegment.pixelStart, top: rows.offset(fillBounds.top) - rowSegment.pixelStart, width: cols.span(fillBounds.left, fillBounds.right), height: rows.span(fillBounds.top, fillBounds.bottom) }} /> : null}
        </div>
        <button aria-label="填充柄" className="infinite-fill-handle" onPointerDown={(event) => { event.preventDefault(); event.stopPropagation(); fillDragRef.current = { source: selection, copy: event.ctrlKey || event.metaKey }; const target = { row: activeBounds.bottom, col: activeBounds.right }; fillTargetRef.current = target; setFillTarget(target); }} style={{ left: cols.offset(activeBounds.right + 1) - colSegment.pixelStart - 4, top: rows.offset(activeBounds.bottom + 1) - rowSegment.pixelStart - 4 }} type="button" />
      </div>
    </div>
  </div>;
}

export default function InfiniteWorkbook() {
  const [identity, setIdentity] = useState<CollaborationIdentity>(() => ({ clientId: createClientId(), name: localStorage.getItem(NAME_KEY) || "用户" }));
  const [sheets, setSheets] = useState<V2SheetSummary[]>([]);
  const [activeSheetId, setActiveSheetId] = useState("");
  const [sheet, setSheet] = useState<SparseSheet | null>(null);
  const [revision, setRevision] = useState(0);
  const [presence, setPresence] = useState<Presence[]>([]);
  const [loading, setLoading] = useState(true);
  const [connection, setConnection] = useState<"connecting" | "live" | "reconnecting">("connecting");
  const [selection, setSelection] = useState<Selection>({ anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } });
  const [extraSelections, setExtraSelections] = useState<Selection[]>([]);
  const [editing, setEditing] = useState<EditingCell>(null);
  const [formulaDraft, setFormulaDraft] = useState("");
  const [formulaFocused, setFormulaFocused] = useState(false);
  const [formulaReferenceStart, setFormulaReferenceStart] = useState<CellPosition | null>(null);
  const [viewportCells, setViewportCells] = useState({ rows: 28, cols: 12 });
  const [navigation, setNavigation] = useState({ rows: 42, cols: 18 });
  const [zoom, setZoom] = useState(100);
  const [toast, setToast] = useState<string | null>(null);
  const [contextMenu, setContextMenu] = useState<{ x: number; y: number; row: number; col: number } | null>(null);
  const [findOpen, setFindOpen] = useState(false);
  const [findQuery, setFindQuery] = useState("");
  const [findCursor, setFindCursor] = useState(0);
  const [formatMenuOpen, setFormatMenuOpen] = useState(false);
  const [fillMenuOpen, setFillMenuOpen] = useState(false);
  const [fontMenuOpen, setFontMenuOpen] = useState(false);
  const [borderEnabled, setBorderEnabled] = useState(false);
  const [formatBrush, setFormatBrush] = useState<CellFormat | null>(null);
  const [formatBrushContinuous, setFormatBrushContinuous] = useState(false);
  const [freeze, setFreeze] = useState(false);
  const [history, setHistory] = useState<HistoryEntry[]>([]);
  const [redoHistory, setRedoHistory] = useState<HistoryEntry[]>([]);
  const activeSheetRef = useRef(activeSheetId);
  const fetchingRef = useRef(new Set<string>());
  // This is deliberately a single nearby cache window, rather than a list of
  // every area the user has browsed.  Blank cells have no objects to evict,
  // and retaining arbitrary old windows would eventually turn navigation
  // through an empty sheet into a hidden full-sheet cache.
  const windowCacheRef = useRef<{ sheetId: string; range: V2UsedRange } | null>(null);
  const latestWindowFetchRef = useRef(0);
  const toastTimerRef = useRef<number | null>(null);
  const formulaInputRef = useRef<HTMLInputElement>(null);
  const formatBrushSourceRef = useRef("");
  const viewportCellsRef = useRef(viewportCells);
  const visibleNavigationRef = useRef<V2UsedRange>({ top: 0, left: 0, bottom: 0, right: 0 });
  const sheetRef = useRef<SparseSheet | null>(sheet);
  activeSheetRef.current = activeSheetId;
  viewportCellsRef.current = viewportCells;
  sheetRef.current = sheet;

  const showToast = useCallback((message: string) => {
    setToast(message);
    if (toastTimerRef.current !== null) window.clearTimeout(toastTimerRef.current);
    toastTimerRef.current = window.setTimeout(() => setToast(null), 2600);
  }, []);

  const initialNavigation = useCallback((content: V2UsedRange, viewport = viewportCellsRef.current) => {
    const bufferRows = Math.max(1, Math.ceil(viewport.rows / 2));
    const bufferCols = Math.max(1, Math.ceil(viewport.cols / 2));
    return {
      rows: Math.max(Math.ceil(viewport.rows * 1.5), content.bottom + 1 + bufferRows),
      cols: Math.max(Math.ceil(viewport.cols * 1.5), content.right + 1 + bufferCols),
    };
  }, []);

  const settleNavigation = useCallback((content: V2UsedRange, visible = visibleNavigationRef.current) => {
    const baseline = initialNavigation(content);
    const bufferRows = Math.max(1, Math.ceil(viewportCellsRef.current.rows / 2));
    const bufferCols = Math.max(1, Math.ceil(viewportCellsRef.current.cols / 2));
    setNavigation((current) => {
      const rows = Math.max(baseline.rows, Math.min(current.rows, Math.max(1, visible.bottom + 1 + bufferRows)));
      const cols = Math.max(baseline.cols, Math.min(current.cols, Math.max(1, visible.right + 1 + bufferCols)));
      return rows === current.rows && cols === current.cols ? current : { rows, cols };
    });
  }, [initialNavigation]);

  const applySnapshot = useCallback((snapshot: V2Snapshot, resetNavigation = false) => {
    setSheets(snapshot.sheets);
    setRevision(snapshot.revision);
    setPresence((snapshot.presence ?? []) as Presence[]);
    setSheet((current) => mergeWindow(current, snapshot));
    if (resetNavigation) setNavigation(initialNavigation(snapshot.sheet.content_range));
  }, [initialNavigation]);

  const loadWindow = useCallback(async (sheetId: string, top = 0, left = 0, rows = 64, cols = 24, resetNavigation = false) => {
    const requested = {
      top: Math.max(0, top),
      left: Math.max(0, left),
      bottom: Math.max(0, top) + Math.max(1, rows) - 1,
      right: Math.max(0, left) + Math.max(1, cols) - 1,
    };
    const cached = windowCacheRef.current;
    if (!resetNavigation && cached?.sheetId === sheetId && windowContains(cached.range, requested)) return;

    // Fetch the active virtual range plus one screen on every side.  The grid
    // reports a new visible range as soon as it crosses a row/column boundary;
    // without this buffer, a slow wheel gesture would produce one request per
    // row.  Keep the request bounded even for a very large viewport.
    const rowBuffer = Math.max(24, viewportCellsRef.current.rows);
    const colBuffer = Math.max(10, viewportCellsRef.current.cols);
    const fetchTop = Math.max(0, requested.top - rowBuffer);
    const fetchLeft = Math.max(0, requested.left - colBuffer);
    const fetchRows = Math.min(2_000, requested.bottom - fetchTop + 1 + rowBuffer);
    const fetchCols = Math.min(2_000, requested.right - fetchLeft + 1 + colBuffer);
    const requestKey = `${sheetId}:${fetchTop}:${fetchLeft}:${fetchRows}:${fetchCols}`;
    if (fetchingRef.current.has(requestKey)) return;
    fetchingRef.current.add(requestKey);
    const fetchNumber = latestWindowFetchRef.current + 1;
    latestWindowFetchRef.current = fetchNumber;
    try {
      const snapshot = await getV2Window(sheetId, fetchTop, fetchLeft, fetchRows, fetchCols);
      // A previous window may finish after a fast scroll.  Its sparse data is
      // still valid, but applying it would prune the newer nearby cache, so
      // discard it instead of letting an old response pull the UI backwards.
      if (activeSheetRef.current !== sheetId || latestWindowFetchRef.current !== fetchNumber) return;
      windowCacheRef.current = { sheetId, range: snapshot.window };
      applySnapshot(snapshot, resetNavigation);
    } catch (error) {
      showToast(error instanceof Error ? error.message : "无法加载工作表");
    } finally {
      fetchingRef.current.delete(requestKey);
      if (activeSheetRef.current === sheetId) setLoading(false);
    }
  }, [applySnapshot, showToast]);

  useEffect(() => {
    let cancelled = false;
    void getV2Sheets().then(({ sheets: currentSheets }) => {
      if (cancelled) return;
      setSheets(currentSheets);
      const remembered = localStorage.getItem(ACTIVE_SHEET_KEY);
      const target = currentSheets.find((item) => item.id === remembered)?.id ?? currentSheets[0]?.id ?? "";
      setActiveSheetId(target);
    }).catch((error) => { if (!cancelled) { showToast(error instanceof Error ? error.message : "无法连接协作工作簿"); setLoading(false); } });
    return () => { cancelled = true; };
  }, [showToast]);

  useEffect(() => {
    if (!activeSheetId) return;
    localStorage.setItem(ACTIVE_SHEET_KEY, activeSheetId);
    setLoading(true);
    setSheet(null);
    setSelection({ anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } });
    setExtraSelections([]);
    setEditing(null);
    setFormulaDraft("");
    setNavigation({ rows: 42, cols: 18 });
    windowCacheRef.current = null;
    latestWindowFetchRef.current += 1;
    void loadWindow(activeSheetId, 0, 0, 64, 24, true);
  }, [activeSheetId, loadWindow]);

  useEffect(() => {
    if (!activeSheetId || !identity.name) return undefined;
    setConnection("connecting");
    const source = new EventSource(v2EventsUrl(identity, activeSheetId));
    source.onopen = () => setConnection("live");
    source.onerror = () => setConnection("reconnecting");
    source.addEventListener("snapshot", (event) => {
      const snapshot = JSON.parse((event as MessageEvent).data) as V2Snapshot;
      if (activeSheetRef.current !== snapshot.sheet.id) return;
      applySnapshot(snapshot);
      setLoading(false);
      setConnection("live");
    });
    source.addEventListener("delta", (event) => {
      const delta = JSON.parse((event as MessageEvent).data) as V2Delta;
      if (activeSheetRef.current !== delta.sheet_id) return;
      setSheet((current) => {
        if (!current || current.id !== delta.sheet_id) return current;
        const cells = { ...current.cells }; const meta = { ...current.meta }; const formats = { ...current.formats };
        delta.cells.forEach((item) => { const key = cellKey(item.row, item.col); if (item.value) cells[key] = item.value; else delete cells[key]; meta[key] = item.meta; });
        delta.formats.forEach((item) => { const key = cellKey(item.row, item.col); if (isMeaningfulFormat(item.format)) formats[key] = item.format; else delete formats[key]; });
        const rowHeights = { ...current.rowHeights, ...delta.row_heights };
        const columnWidths = { ...current.columnWidths, ...delta.column_widths };
        Object.entries(rowHeights).forEach(([key, value]) => { if (value === DEFAULT_ROW_HEIGHT) delete rowHeights[key]; });
        Object.entries(columnWidths).forEach(([key, value]) => { if (value === DEFAULT_COLUMN_WIDTH) delete columnWidths[key]; });
        return { ...current, cells, meta, formats, rowHeights, columnWidths, merges: delta.merges ?? current.merges, usedRange: delta.used_range, contentRange: delta.content_range };
      });
      setRevision(delta.revision);
    });
    source.addEventListener("presence", (event) => {
      const payload = JSON.parse((event as MessageEvent).data) as { sheet_id: string; presence: Presence[] };
      if (payload.sheet_id === activeSheetRef.current) setPresence(payload.presence);
    });
    source.addEventListener("workbook", (event) => {
      const payload = JSON.parse((event as MessageEvent).data) as { sheets: V2SheetSummary[] };
      setSheets(payload.sheets);
      if (!payload.sheets.some((item) => item.id === activeSheetRef.current)) setActiveSheetId(payload.sheets[0]?.id ?? "");
    });
    return () => source.close();
  }, [activeSheetId, applySnapshot, identity]);

  useEffect(() => {
    if (!activeSheetId || !identity.name) return undefined;
    const send = () => void updateV2Presence(identity, activeSheetId, selection.focus).catch(() => undefined);
    send();
    const interval = window.setInterval(send, 20_000);
    return () => window.clearInterval(interval);
  }, [activeSheetId, identity, selection.focus.col, selection.focus.row]);

  useEffect(() => () => { if (toastTimerRef.current !== null) window.clearTimeout(toastTimerRef.current); }, []);

  useEffect(() => {
    if (!sheet) return;
    settleNavigation(sheet.contentRange);
  }, [settleNavigation, sheet?.contentRange.bottom, sheet?.contentRange.right, sheet?.id]);

  const fetchVisibleWindow = useCallback((top: number, left: number, rows: number, cols: number) => {
    if (!activeSheetId) return;
    void loadWindow(activeSheetId, top, left, rows, cols);
  }, [activeSheetId, loadWindow]);

  const onViewport = useCallback((rows: number, cols: number) => setViewportCells((current) => current.rows === rows && current.cols === cols ? current : { rows, cols }), []);

  const onVisibleRange = useCallback((range: V2UsedRange) => {
    visibleNavigationRef.current = range;
    const current = sheetRef.current;
    if (current) settleNavigation(current.contentRange, range);
  }, [settleNavigation]);

  const ensureNavigationFor = useCallback((target: CellPosition | Bounds) => {
    const bounds = "top" in target ? target : { top: target.row, bottom: target.row, left: target.col, right: target.col };
    const rowBuffer = Math.max(viewportCellsRef.current.rows, 24);
    const colBuffer = Math.max(viewportCellsRef.current.cols, 10);
    setNavigation((current) => {
      const rows = Math.max(current.rows, bounds.bottom + 1 + rowBuffer);
      const cols = Math.max(current.cols, bounds.right + 1 + colBuffer);
      return rows === current.rows && cols === current.cols ? current : { rows, cols };
    });
  }, []);

  const onWheelExtend = useCallback((axis: "row" | "col") => {
    // Scrollbar drag never calls this.  This is intentionally reachable only
    // from the wheel handler at the native track's current endpoint.
    setNavigation((current) => axis === "row"
      ? { ...current, rows: current.rows + Math.max(viewportCells.rows, 24) }
      : { ...current, cols: current.cols + Math.max(viewportCells.cols, 10) });
  }, [viewportCells.cols, viewportCells.rows]);

  const onFillExtend = useCallback((axis: "row" | "col") => {
    setNavigation((current) => axis === "row"
      ? { ...current, rows: current.rows + Math.max(viewportCellsRef.current.rows, 24) }
      : { ...current, cols: current.cols + Math.max(viewportCellsRef.current.cols, 10) });
  }, []);

  const selectCell = useCallback((row: number, col: number, extend: boolean, additive: boolean, grow = false) => {
    const next = { row: Math.max(0, row), col: Math.max(0, col) };
    if (grow) ensureNavigationFor(next);
    setSelection((current) => ({ anchor: extend ? current.anchor : next, focus: next }));
    setExtraSelections((current) => extend ? current : additive ? [...current, selection] : []);
    setContextMenu(null);
  }, [ensureNavigationFor, selection]);

  const dragSelection = useCallback((row: number, col: number) => setSelection((current) => ({ ...current, focus: { row, col } })), []);

  const selectRow = useCallback((row: number, extend: boolean, additive: boolean) => {
    setSelection((current) => ({ anchor: extend ? current.anchor : { row, col: 0 }, focus: { row, col: Math.max(0, navigation.cols - 1) } }));
    setExtraSelections((current) => extend ? current : additive ? [...current, selection] : []);
  }, [navigation.cols, selection]);

  const selectColumn = useCallback((col: number, extend: boolean, additive: boolean) => {
    setSelection((current) => ({ anchor: extend ? current.anchor : { row: 0, col }, focus: { row: Math.max(0, navigation.rows - 1), col } }));
    setExtraSelections((current) => extend ? current : additive ? [...current, selection] : []);
  }, [navigation.rows, selection]);

  const selectAll = useCallback(() => {
    setExtraSelections([]);
    setSelection({ anchor: { row: 0, col: 0 }, focus: { row: Math.max(0, navigation.rows - 1), col: Math.max(0, navigation.cols - 1) } });
  }, [navigation.cols, navigation.rows]);

  const activeBounds = boundsOf(selection);
  const activeKey = cellKey(selection.focus.row, selection.focus.col);
  const activeValue = sheet?.cells[activeKey] ?? "";
  const activeFormat = sheet?.formats[activeKey] ?? {};
  const allSelections = useMemo(() => [...extraSelections, selection], [extraSelections, selection]);

  useEffect(() => {
    if (!editing && !formulaFocused) setFormulaDraft(activeValue);
  }, [activeKey, activeValue, editing, formulaFocused]);

  const selectedPositions = useCallback((limit = MAX_BATCH_CELLS): CellPosition[] => {
    const output: CellPosition[] = [];
    for (const item of allSelections) {
      const bounds = boundsOf(item);
      for (let row = bounds.top; row <= bounds.bottom; row += 1) for (let col = bounds.left; col <= bounds.right; col += 1) {
        output.push({ row, col });
        if (output.length > limit) throw new Error(`选区超过 ${limit.toLocaleString()} 个单元格，请缩小后再操作`);
      }
    }
    return Array.from(new Map(output.map((item) => [cellKey(item.row, item.col), item])).values());
  }, [allSelections]);

  const applyDelta = useCallback((delta: V2Delta) => {
    if (delta.sheet_id !== activeSheetRef.current) return;
    setSheet((current) => {
      if (!current || current.id !== delta.sheet_id) return current;
      const cells = { ...current.cells }; const meta = { ...current.meta }; const formats = { ...current.formats };
      delta.cells.forEach((item) => { const key = cellKey(item.row, item.col); if (item.value) cells[key] = item.value; else delete cells[key]; meta[key] = item.meta; });
      delta.formats.forEach((item) => { const key = cellKey(item.row, item.col); if (isMeaningfulFormat(item.format)) formats[key] = item.format; else delete formats[key]; });
      const rowHeights = { ...current.rowHeights, ...delta.row_heights }; const columnWidths = { ...current.columnWidths, ...delta.column_widths };
      Object.entries(rowHeights).forEach(([key, value]) => { if (value === DEFAULT_ROW_HEIGHT) delete rowHeights[key]; });
      Object.entries(columnWidths).forEach(([key, value]) => { if (value === DEFAULT_COLUMN_WIDTH) delete columnWidths[key]; });
      return { ...current, cells, meta, formats, rowHeights, columnWidths, merges: delta.merges ?? current.merges, usedRange: delta.used_range, contentRange: delta.content_range };
    });
    setRevision(delta.revision);
  }, []);

  const submit = useCallback(async (operations: Array<Record<string, unknown>>, message?: string, historyEntry?: HistoryEntry) => {
    if (!activeSheetId || !sheet) return;
    try {
      const delta = await applyV2Operations(identity, activeSheetId, operations, selection.focus);
      applyDelta(delta);
      if (historyEntry) { setHistory((current) => [...current.slice(-99), historyEntry]); setRedoHistory([]); }
      if (message) showToast(message);
    } catch (error) {
      showToast(error instanceof Error ? error.message : "保存失败");
      void loadWindow(activeSheetId, Math.max(0, selection.focus.row - 80), Math.max(0, selection.focus.col - 30), 180, 80);
    }
  }, [activeSheetId, applyDelta, identity, loadWindow, selection.focus, sheet, showToast]);

  const beginEdit = useCallback((row: number, col: number, seed?: string) => {
    setFormulaReferenceStart(null);
    setEditing({ row, col, value: seed ?? (sheet?.cells[cellKey(row, col)] ?? "") });
    setFormulaDraft(seed ?? (sheet?.cells[cellKey(row, col)] ?? ""));
  }, [sheet?.cells]);

  const commitEdit = useCallback(() => {
    if (!editing || !sheet) return;
    const before = sheet.cells[cellKey(editing.row, editing.col)] ?? "";
    const next = editing.value;
    setEditing(null);
    if (before === next) return;
    void submit([{ type: "cells", cells: [{ row: editing.row, col: editing.col, value: next, version: sheet.meta[cellKey(editing.row, editing.col)]?.version ?? 0 }] }], "已保存", { undo: [{ type: "cells", cells: [{ row: editing.row, col: editing.col, value: before }] }], redo: [{ type: "cells", cells: [{ row: editing.row, col: editing.col, value: next }] }] });
  }, [editing, sheet, submit]);

  const commitFormula = useCallback(() => {
    if (!sheet) return;
    const row = selection.focus.row; const col = selection.focus.col;
    const before = sheet.cells[cellKey(row, col)] ?? "";
    const next = formulaDraft;
    setFormulaFocused(false);
    if (before === next) return;
    void submit([{ type: "cells", cells: [{ row, col, value: next, version: sheet.meta[cellKey(row, col)]?.version ?? 0 }] }], "已保存", { undo: [{ type: "cells", cells: [{ row, col, value: before }] }], redo: [{ type: "cells", cells: [{ row, col, value: next }] }] });
  }, [formulaDraft, selection.focus.col, selection.focus.row, sheet, submit]);

  const copySelection = useCallback(async () => {
    if (!sheet) return;
    const bounds = boundsOf(selection);
    const values = Array.from({ length: bounds.bottom - bounds.top + 1 }, (_, rowOffset) => Array.from({ length: bounds.right - bounds.left + 1 }, (_, colOffset) => sheet.cells[cellKey(bounds.top + rowOffset, bounds.left + colOffset)] ?? ""));
    const text = serializeSpreadsheetClipboard(values);
    try { await navigator.clipboard.writeText(text); showToast("已复制"); }
    catch { window.prompt("复制表格内容", text); }
  }, [selection, sheet, showToast]);

  const applyValueMatrix = useCallback(async (matrix: string[][], label: string) => {
    if (!sheet || !matrix.length) return;
    const start = selection.focus;
    const cells: Array<Record<string, unknown>> = [];
    const before: Array<Record<string, unknown>> = [];
    const redo: Array<Record<string, unknown>> = [];
    matrix.forEach((line, rowOffset) => line.forEach((value, colOffset) => {
      const row = start.row + rowOffset; const col = start.col + colOffset; const key = cellKey(row, col);
      cells.push({ row, col, value, version: sheet.meta[key]?.version ?? 0 });
      before.push({ row, col, value: sheet.cells[key] ?? "" });
      redo.push({ row, col, value });
    }));
    if (cells.length > MAX_BATCH_CELLS) { showToast(`一次最多粘贴 ${MAX_BATCH_CELLS.toLocaleString()} 个单元格`); return; }
    const end = { row: start.row + matrix.length - 1, col: start.col + Math.max(...matrix.map((line) => line.length), 1) - 1 };
    ensureNavigationFor({ top: start.row, bottom: end.row, left: start.col, right: end.col });
    setSelection({ anchor: start, focus: end });
    await submit([{ type: "cells", cells }], `${label} ${cells.length} 个单元格`, { undo: [{ type: "cells", cells: before }], redo: [{ type: "cells", cells: redo }] });
  }, [ensureNavigationFor, selection.focus, sheet, showToast, submit]);

  const pasteText = useCallback(async (text: string) => { await applyValueMatrix(parseSpreadsheetClipboard(text), "已粘贴"); }, [applyValueMatrix]);

  const pasteFromClipboard = useCallback(async () => {
    setContextMenu(null);
    try {
      if (navigator.clipboard.read) {
        const items = await navigator.clipboard.read();
        for (const item of items) if (item.types.includes("text/html")) {
          const matrix = parseHtmlSpreadsheetClipboard(await (await item.getType("text/html")).text());
          if (matrix.length) { await applyValueMatrix(matrix, "已粘贴"); return; }
        }
      }
      await pasteText(await navigator.clipboard.readText());
    } catch { showToast("浏览器未允许读取剪贴板，请直接按 Ctrl+V"); }
  }, [applyValueMatrix, pasteText, showToast]);

  const clearSelection = useCallback(async () => {
    if (!sheet) return;
    let positions: CellPosition[];
    try { positions = selectedPositions(); } catch (error) { showToast(error instanceof Error ? error.message : "选区过大"); return; }
    const cells = positions.map(({ row, col }) => ({ row, col, value: "", version: sheet.meta[cellKey(row, col)]?.version ?? 0 }));
    const before = positions.map(({ row, col }) => ({ row, col, value: sheet.cells[cellKey(row, col)] ?? "" }));
    await submit([{ type: "cells", cells }], `已清空 ${cells.length} 个单元格`, { undo: [{ type: "cells", cells: before }], redo: [{ type: "cells", cells: positions.map(({ row, col }) => ({ row, col, value: "" })) }] });
  }, [selectedPositions, sheet, showToast, submit]);

  const cutSelection = useCallback(async () => { await copySelection(); await clearSelection(); }, [clearSelection, copySelection]);

  const applyFormat = useCallback(async (patch: CellFormat) => {
    let positions: CellPosition[];
    try { positions = selectedPositions(); } catch (error) { showToast(error instanceof Error ? error.message : "选区过大"); return; }
    if (!positions.length || !sheet) return;
    const before = positions.map(({ row, col }) => ({ row, col, format: sheet.formats[cellKey(row, col)] ?? {} }));
    const after = positions.map(({ row, col }) => ({ row, col, format: { ...(sheet.formats[cellKey(row, col)] ?? {}), ...patch } }));
    await submit([{ type: "format", cells: positions, format: patch }], "已应用单元格格式", { undo: [{ type: "format", cells: before.map(({ row, col }) => ({ row, col })), format: {} }], redo: [{ type: "format", cells: after.map(({ row, col }) => ({ row, col })), format: patch }] });
  }, [selectedPositions, sheet, showToast, submit]);

  const applyBorders = useCallback(async () => {
    const next = !borderEnabled;
    setBorderEnabled(next);
    const edge = { style: "solid" as const, width: 1 as const, color: "#66717d" };
    await applyFormat(next ? { border: true, borders: { top: edge, right: edge, bottom: edge, left: edge } } : { border: false, borders: {} });
  }, [applyFormat, borderEnabled]);

  const setRowHeight = useCallback((row: number, height: number, persist = true) => {
    setSheet((current) => current ? { ...current, rowHeights: height === DEFAULT_ROW_HEIGHT ? Object.fromEntries(Object.entries(current.rowHeights).filter(([key]) => key !== String(row))) : { ...current.rowHeights, [row]: height } } : current);
    if (persist) void submit([{ type: "layout", action: "row_height", payload: { row, height } }]);
  }, [submit]);

  const setColumnWidth = useCallback((col: number, width: number, persist = true) => {
    setSheet((current) => current ? { ...current, columnWidths: width === DEFAULT_COLUMN_WIDTH ? Object.fromEntries(Object.entries(current.columnWidths).filter(([key]) => key !== String(col))) : { ...current.columnWidths, [col]: width } } : current);
    if (persist) void submit([{ type: "layout", action: "column_width", payload: { col, width } }]);
  }, [submit]);

  const mergedInSelection = useMemo(() => {
    if (!sheet) return [];
    const bounds = boundsOf(selection);
    return sheet.merges.filter((merge) => !(merge.end.row < bounds.top || merge.start.row > bounds.bottom || merge.end.col < bounds.left || merge.start.col > bounds.right));
  }, [selection, sheet]);

  const mergeSelection = useCallback(() => {
    const bounds = boundsOf(selection);
    if (bounds.top === bounds.bottom && bounds.left === bounds.right) { showToast("请先选择两个或更多单元格"); return; }
    void submit([{ type: "layout", action: "merge", payload: { start: { row: bounds.top, col: bounds.left }, end: { row: bounds.bottom, col: bounds.right } } }], "已合并单元格");
  }, [selection, showToast, submit]);

  const unmergeSelection = useCallback(() => {
    if (!mergedInSelection.length) { showToast("选区中没有合并单元格"); return; }
    void submit([{ type: "layout", action: "unmerge", payload: { ids: mergedInSelection.map((merge) => merge.id) } }], "已取消合并");
  }, [mergedInSelection, showToast, submit]);

  const fillSelection = useCallback((source: Selection, target: CellPosition, copy: boolean) => {
    if (!sheet) return;
    const origin = boundsOf(source);
    const bounds = { top: Math.min(origin.top, target.row), bottom: Math.max(origin.bottom, target.row), left: Math.min(origin.left, target.col), right: Math.max(origin.right, target.col) };
    ensureNavigationFor(bounds);
    const updates: Array<Record<string, unknown>> = [];
    const sourceRows = origin.bottom - origin.top + 1; const sourceCols = origin.right - origin.left + 1;
    for (let row = bounds.top; row <= bounds.bottom; row += 1) for (let col = bounds.left; col <= bounds.right; col += 1) {
      if (row >= origin.top && row <= origin.bottom && col >= origin.left && col <= origin.right) continue;
      const sourceRow = origin.top + ((row - origin.top) % sourceRows + sourceRows) % sourceRows;
      const sourceCol = origin.left + ((col - origin.left) % sourceCols + sourceCols) % sourceCols;
      const raw = sheet.cells[cellKey(sourceRow, sourceCol)] ?? "";
      let value = raw;
      if (!copy && /^\s*=/.test(raw)) value = translateFormula(raw, row - sourceRow, col - sourceCol);
      else if (!copy && sourceRows === 1 && sourceCols === 1 && /^-?\d+(?:\.\d+)?$/.test(raw)) {
        const offset = Math.abs(row - origin.top) >= Math.abs(col - origin.left) ? row - origin.top : col - origin.left;
        value = String(Number(raw) + offset);
      }
      updates.push({ row, col, value, version: sheet.meta[cellKey(row, col)]?.version ?? 0 });
    }
    if (!updates.length) return;
    void submit([{ type: "cells", cells: updates }], `已填充 ${updates.length} 个单元格`);
  }, [ensureNavigationFor, sheet, submit]);

  const insertOrDelete = useCallback((kind: "insert_rows" | "delete_rows" | "insert_columns" | "delete_columns") => {
    const bounds = boundsOf(selection);
    const rows = kind.endsWith("rows");
    const at = rows ? bounds.top : bounds.left;
    const count = rows ? bounds.bottom - bounds.top + 1 : bounds.right - bounds.left + 1;
    void submit([{ type: kind, at, count }], kind.startsWith("insert") ? "已插入" : "已删除");
  }, [selection, submit]);

  const undo = useCallback(() => {
    const entry = history.at(-1);
    if (!entry) { showToast("没有可撤销的操作"); return; }
    setHistory((current) => current.slice(0, -1));
    setRedoHistory((current) => [...current, entry]);
    void submit(entry.undo, "已撤销");
  }, [history, showToast, submit]);

  const redo = useCallback(() => {
    const entry = redoHistory.at(-1);
    if (!entry) { showToast("没有可重做的操作"); return; }
    setRedoHistory((current) => current.slice(0, -1));
    setHistory((current) => [...current, entry]);
    void submit(entry.redo, "已重做");
  }, [redoHistory, showToast, submit]);

  const selectFormulaReference = useCallback((row: number, col: number, extend: boolean) => {
    const current = editing?.value ?? (formulaFocused ? formulaDraft : "");
    if (!current.trimStart().startsWith("=") || !current.includes("(")) return false;
    const start = extend && formulaReferenceStart ? formulaReferenceStart : { row, col };
    setFormulaReferenceStart(start);
    const reference = `${columnName(start.col)}${start.row + 1}${start.row === row && start.col === col ? "" : `:${columnName(col)}${row + 1}`}`;
    const opener = current.lastIndexOf("(");
    const before = current.slice(0, opener + 1);
    const existing = current.slice(opener + 1).replace(/\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?$/, "");
    const next = `${before}${existing}${reference}${current.endsWith(")") ? ")" : ""}`;
    if (editing) setEditing({ ...editing, value: next });
    else setFormulaDraft(next);
    return true;
  }, [editing, formulaDraft, formulaFocused, formulaReferenceStart]);

  const chooseFormula = useCallback((name: string) => {
    const current = formulaDraft || "=";
    const input = formulaInputRef.current;
    const cursor = input?.selectionStart ?? current.length;
    const token = current.slice(0, cursor).match(/[A-Za-z]*$/)?.[0] ?? "";
    const prefix = current.slice(0, cursor - token.length);
    const suffix = current.slice(input?.selectionEnd ?? cursor);
    const next = `${prefix}${name}()${suffix}`;
    setFormulaDraft(next);
    setFormulaFocused(true);
    requestAnimationFrame(() => {
      formulaInputRef.current?.focus();
      const position = prefix.length + name.length + 1;
      formulaInputRef.current?.setSelectionRange(position, position);
    });
  }, [formulaDraft]);

  const formulaSuggestions = useMemo(() => {
    const match = formulaDraft.trimStart().match(/^=\s*([A-Za-z]*)$/);
    if (!match) return [];
    const keyword = match[1].toUpperCase();
    return FORMULA_NAMES.filter((name) => !keyword || name.startsWith(keyword));
  }, [formulaDraft]);

  const findMatches = useMemo(() => {
    const keyword = findQuery.trim().toLocaleLowerCase();
    if (!keyword || !sheet) return [];
    return Object.entries(sheet.cells).filter(([, value]) => value.toLocaleLowerCase().includes(keyword)).map(([key]) => parseCoordinate(key));
  }, [findQuery, sheet]);

  const moveFind = useCallback((direction: number) => {
    if (!findMatches.length) return;
    setFindCursor((current) => {
      const next = (current + direction + findMatches.length) % findMatches.length;
      const item = findMatches[next];
      setSelection({ anchor: item, focus: item });
      return next;
    });
  }, [findMatches]);

  const renameActiveSheet = useCallback(async () => {
    const current = sheets.find((item) => item.id === activeSheetId);
    const next = window.prompt("重命名工作表", current?.name ?? "");
    if (!next?.trim()) return;
    try {
      const response = await renameV2Sheet(identity, activeSheetId, next.trim());
      setSheets(response.sheets);
      setSheet((currentSheet) => currentSheet ? { ...currentSheet, name: next.trim() } : currentSheet);
      showToast("工作表已重命名");
    } catch (error) { showToast(error instanceof Error ? error.message : "重命名失败"); }
  }, [activeSheetId, identity, sheets, showToast]);

  const addSheet = useCallback(async () => {
    const number = sheets.length + 1;
    let name = `工作表${number}`;
    while (sheets.some((item) => item.name === name)) name = `工作表${Number(name.replace("工作表", "")) + 1}`;
    try {
      const response = await createV2Sheet(identity, name);
      setSheets(response.sheets);
      setActiveSheetId(response.sheet.id);
      showToast(`${name} 已创建`);
    } catch (error) { showToast(error instanceof Error ? error.message : "创建工作表失败"); }
  }, [identity, sheets, showToast]);

  const removeActiveSheet = useCallback(async () => {
    if (sheets.length <= 1) { showToast("至少保留一个工作表"); return; }
    const current = sheets.find((item) => item.id === activeSheetId);
    if (!window.confirm(`确定删除工作表“${current?.name ?? ""}”吗？`)) return;
    try {
      const response = await deleteV2Sheet(identity, activeSheetId);
      setSheets(response.sheets);
      setActiveSheetId(response.sheets[0]?.id ?? "");
      showToast("工作表已删除");
    } catch (error) { showToast(error instanceof Error ? error.message : "删除工作表失败"); }
  }, [activeSheetId, identity, sheets, showToast]);

  const statistics = useMemo(() => {
    if (!sheet) return { count: 0, sum: 0, average: 0 };
    let positions: CellPosition[];
    try { positions = selectedPositions(MAX_BATCH_CELLS); }
    catch { return { count: 0, sum: 0, average: 0 }; }
    const values = positions.map(({ row, col }) => sheet.cells[cellKey(row, col)] ?? "");
    const numbers = values.map((value) => Number(value.replaceAll(",", ""))).filter((value) => Number.isFinite(value));
    const sum = numbers.reduce((total, value) => total + value, 0);
    return { count: values.filter((value) => value !== "").length, sum, average: numbers.length ? sum / numbers.length : 0 };
  }, [selectedPositions, sheet]);

  const onGridContext = useCallback((event: ReactMouseEvent, row: number, col: number) => {
    event.preventDefault();
    const bounds = boundsOf(selection);
    if (row < bounds.top || row > bounds.bottom || col < bounds.left || col > bounds.right) selectCell(row, col, false, false);
    setContextMenu({ x: event.clientX, y: event.clientY, row, col });
  }, [selectCell, selection]);

  const handleTableKeyDown = (event: KeyboardEvent<HTMLDivElement>) => {
    const target = event.target as HTMLElement;
    if (target.matches("input, textarea")) return;
    const meta = event.ctrlKey || event.metaKey;
    if (meta && event.key.toLowerCase() === "c") { event.preventDefault(); void copySelection(); return; }
    if (meta && event.key.toLowerCase() === "x") { event.preventDefault(); void cutSelection(); return; }
    if (meta && event.key.toLowerCase() === "v") return;
    if (meta && event.key.toLowerCase() === "b") { event.preventDefault(); void applyFormat({ bold: !activeFormat.bold }); return; }
    if (meta && event.key.toLowerCase() === "i") { event.preventDefault(); void applyFormat({ italic: !activeFormat.italic }); return; }
    if (meta && event.key.toLowerCase() === "u") { event.preventDefault(); void applyFormat({ underline: !activeFormat.underline }); return; }
    if (event.key === "Delete" || event.key === "Backspace") { event.preventDefault(); void clearSelection(); return; }
    const movement: Record<string, CellPosition> = { ArrowUp: { row: -1, col: 0 }, ArrowDown: { row: 1, col: 0 }, ArrowLeft: { row: 0, col: -1 }, ArrowRight: { row: 0, col: 1 }, Tab: { row: 0, col: 1 }, Enter: { row: 1, col: 0 } };
    const move = movement[event.key];
    if (move) {
      event.preventDefault();
      const row = Math.max(0, selection.focus.row + move.row);
      const col = Math.max(0, selection.focus.col + move.col);
      selectCell(row, col, event.shiftKey, false, true);
      return;
    }
    if (event.key === "F2") { event.preventDefault(); beginEdit(selection.focus.row, selection.focus.col); return; }
    if (event.key.length === 1 && !meta && !event.altKey) { event.preventDefault(); beginEdit(selection.focus.row, selection.focus.col, event.key); }
  };

  const handlePaste = (event: ClipboardEvent<HTMLDivElement>) => {
    if ((event.target as HTMLElement).matches("input, textarea")) return;
    event.preventDefault();
    const html = event.clipboardData.getData("text/html");
    const matrix = html ? parseHtmlSpreadsheetClipboard(html) : [];
    void applyValueMatrix(matrix.length ? matrix : parseSpreadsheetClipboard(event.clipboardData.getData("text/plain")), "已粘贴");
  };

  useEffect(() => {
    if (!formatBrush) return;
    const target = cellKey(selection.focus.row, selection.focus.col);
    if (target === formatBrushSourceRef.current) return;
    void applyFormat(formatBrush);
    if (!formatBrushContinuous) setFormatBrush(null);
  }, [applyFormat, formatBrush, formatBrushContinuous, selection.focus.col, selection.focus.row]);

  if (loading || !sheet) return <div className="infinite-workbook-loading">正在打开无限协作表格…</div>;

  return <div className="app-shell excel-app-shell infinite-workbook-app" onContextMenu={(event) => event.preventDefault()} onMouseDown={() => setContextMenu(null)}>
    <header className="topbar">
      <div className="brand-block"><div className="brand-mark" aria-hidden="true"><span /><span /><span /><span /></div><h1>空白表格</h1><span className="sheet-save-state">{connection === "live" ? "所有编辑内容都会自动保存到云端" : connection === "reconnecting" ? "正在重新连接…" : "正在连接…"}</span></div>
      <div className="topbar-right"><button className="identity-button" onClick={() => { const name = window.prompt("显示名称", identity.name); if (name?.trim()) { localStorage.setItem(NAME_KEY, name.trim()); setIdentity((current) => ({ ...current, name: name.trim() })); } }} type="button">{identity.name}</button></div>
    </header>
    <main className="workspace excel-workspace"><section className="workspace-main">
      <div className="spreadsheet-toolbar excel-toolbar infinite-toolbar" aria-label="表格工具栏">
        <div className="toolbar-group">
          <button aria-label="撤销" onClick={undo} type="button"><Undo2 /></button><button aria-label="重做" onClick={redo} type="button"><Redo2 /></button>
          <button aria-label="格式刷" className={formatBrush ? "active" : ""} onClick={(event) => { if (formatBrush && event.detail > 1) { setFormatBrushContinuous(true); showToast("连续格式刷已开启"); return; } if (formatBrush) { setFormatBrush(null); setFormatBrushContinuous(false); return; } formatBrushSourceRef.current = activeKey; setFormatBrush(activeFormat); setFormatBrushContinuous(false); showToast("请选择目标区域应用格式"); }} type="button"><Paintbrush />格式刷</button>
          <button aria-label="复制" onClick={() => void copySelection()} type="button"><Copy />复制</button><button aria-label="剪切" onClick={() => void cutSelection()} type="button"><Scissors />剪切</button><button aria-label="粘贴" onClick={() => void pasteFromClipboard()} type="button"><ClipboardPaste />粘贴</button>
        </div>
        <div className="toolbar-group formatting-group">
          <select aria-label="字体" onChange={(event) => void applyFormat({ font_family: event.target.value })} value={activeFormat.font_family ?? "Microsoft YaHei"}><option value="Microsoft YaHei">微软雅黑</option><option value="Arial">Arial</option><option value="SimSun">宋体</option><option value="KaiTi">楷体</option></select>
          <select aria-label="字号" onChange={(event) => void applyFormat({ font_size: Number(event.target.value) })} value={activeFormat.font_size ?? 12}>{[9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36].map((size) => <option key={size} value={size}>{size}</option>)}</select>
          <button aria-label="加粗" className={activeFormat.bold ? "active" : ""} onClick={() => void applyFormat({ bold: !activeFormat.bold })} type="button"><b>B</b></button><button aria-label="斜体" className={activeFormat.italic ? "active" : ""} onClick={() => void applyFormat({ italic: !activeFormat.italic })} type="button"><i>I</i></button><button aria-label="下划线" className={activeFormat.underline ? "active" : ""} onClick={() => void applyFormat({ underline: !activeFormat.underline })} type="button"><u>U</u></button>
          <button aria-label="框线" className={borderEnabled ? "active" : ""} onClick={() => void applyBorders()} type="button"><Grid3X3 />框线</button>
          <span className="toolbar-menu-anchor"><button aria-label="填充" className={fillMenuOpen ? "active" : ""} onClick={() => { setFillMenuOpen((current) => !current); setFontMenuOpen(false); }} type="button"><PaintBucket />填充</button>{fillMenuOpen ? <span className="infinite-color-menu">{FILL_COLORS.map((color) => <button aria-label={color} key={color} onClick={() => { void applyFormat({ fill: color }); setFillMenuOpen(false); }} style={{ background: color }} type="button" />)}</span> : null}</span>
          <span className="toolbar-menu-anchor"><button aria-label="字色" className={fontMenuOpen ? "active" : ""} onClick={() => { setFontMenuOpen((current) => !current); setFillMenuOpen(false); }} type="button"><Palette />字色</button>{fontMenuOpen ? <span className="infinite-color-menu">{FONT_COLORS.map((color) => <button aria-label={color} key={color} onClick={() => { void applyFormat({ font_color: color }); setFontMenuOpen(false); }} style={{ background: color }} type="button" />)}</span> : null}</span>
          <button aria-label="清除" onClick={() => void clearSelection()} type="button"><Eraser />清除</button>
        </div>
        <div className="toolbar-group formatting-group">
          <button aria-label="顶端对齐" onClick={() => void applyFormat({ vertical_align: "top" })} type="button"><AlignStartVertical /></button><button aria-label="垂直居中" onClick={() => void applyFormat({ vertical_align: "middle" })} type="button"><AlignCenterVertical /></button><button aria-label="底端对齐" onClick={() => void applyFormat({ vertical_align: "bottom" })} type="button"><AlignEndVertical /></button>
          <button aria-label="换行" className={activeFormat.wrap ? "active" : ""} onClick={() => void applyFormat({ wrap: !activeFormat.wrap })} type="button"><WrapText /></button><button aria-label="左对齐" onClick={() => void applyFormat({ align: "left" })} type="button"><AlignLeft /></button><button aria-label="居中对齐" onClick={() => void applyFormat({ align: "center" })} type="button"><AlignCenter /></button><button aria-label="右对齐" onClick={() => void applyFormat({ align: "right" })} type="button"><AlignRight /></button><button aria-label="两端对齐" onClick={() => void applyFormat({ align: "justify" })} type="button"><AlignJustify /></button>
        </div>
        <div className="toolbar-group">
          <button aria-label={mergedInSelection.length ? "拆分单元格" : "合并单元格"} onClick={mergedInSelection.length ? unmergeSelection : mergeSelection} type="button">{mergedInSelection.length ? <TableCellsSplit /> : <TableCellsMerge />}合并</button>
          <button aria-label="格式" className={formatMenuOpen ? "active" : ""} onClick={() => setFormatMenuOpen((current) => !current)} type="button"><TableProperties />格式</button>
          <button aria-label="冻结表头" className={freeze ? "active" : ""} onClick={() => setFreeze((current) => !current)} type="button"><Snowflake /></button>
          <button aria-label="查找" className={findOpen ? "active" : ""} onClick={() => { setFindOpen((current) => !current); setFindCursor(0); }} type="button"><Search /></button>
        </div>
      </div>
      {formatMenuOpen ? <div className="format-panel infinite-format-panel" role="dialog"><div><strong>单元格格式</strong><button aria-label="关闭" onClick={() => setFormatMenuOpen(false)} type="button"><X /></button></div>{[["general", "常规"], ["number", "数值"], ["date", "日期"], ["percentage", "百分比"], ["text", "文本"]].map(([value, label]) => <button className={(activeFormat.number_format ?? "general") === value ? "active" : ""} key={value} onClick={() => { void applyFormat({ number_format: value as NonNullable<CellFormat["number_format"]> }); setFormatMenuOpen(false); }} type="button">{label}</button>)}</div> : null}
      {findOpen ? <div className="find-panel infinite-find-panel" role="dialog"><input autoFocus aria-label="查找" onChange={(event) => { setFindQuery(event.target.value); setFindCursor(0); }} placeholder="查找" value={findQuery} /><span>{findQuery ? `${findMatches.length} 项匹配` : ""}</span><button onClick={() => moveFind(-1)} type="button">上一个</button><button onClick={() => moveFind(1)} type="button">下一个</button><button aria-label="关闭查找" className="panel-close" onClick={() => setFindOpen(false)} type="button"><X /></button></div> : null}
      <div className="formula-bar infinite-formula-bar"><div className="formula-name-card"><input aria-label="名称框" className="formula-name-box" readOnly value={`${columnName(selection.focus.col)}${selection.focus.row + 1}`} /></div><div className="formula-editor-card"><span className="formula-fx" aria-hidden="true">fx</span><input aria-label="单元格内容" className="formula-content-input" onBlur={() => { if (formulaFocused) commitFormula(); }} onChange={(event) => setFormulaDraft(event.target.value)} onFocus={() => { setFormulaFocused(true); setFormulaDraft(activeValue); }} onKeyDown={(event) => { if (event.key === "Enter") { event.preventDefault(); commitFormula(); } if (event.key === "Escape") { event.preventDefault(); setFormulaFocused(false); setFormulaDraft(activeValue); } }} ref={formulaInputRef} value={formulaFocused ? formulaDraft : activeValue} />{formulaFocused && formulaSuggestions.length ? <div className="formula-suggestions">{formulaSuggestions.map((name) => <button key={name} onMouseDown={(event) => { event.preventDefault(); chooseFormula(name); }} type="button"><span>fx</span>{name}</button>)}</div> : null}</div></div>
      <section className={`table-card excel-table-card infinite-table-card${freeze ? " freeze-active" : ""}`} onCopy={(event) => { event.preventDefault(); void copySelection(); }} onCut={(event) => { event.preventDefault(); void cutSelection(); }} onKeyDown={handleTableKeyDown} onPaste={handlePaste} tabIndex={0}>
        <InfiniteGrid editing={editing} formulaDraft={editing?.value ?? (formulaFocused ? formulaDraft : "")} identity={identity} navigation={navigation} onBeginEdit={beginEdit} onCancelEdit={() => setEditing(null)} onCommitEdit={commitEdit} onContext={onGridContext} onDragSelect={dragSelection} onEditChange={(value) => setEditing((current) => current ? { ...current, value } : current)} onFetchWindow={fetchVisibleWindow} onFill={fillSelection} onFillExtend={onFillExtend} onFormulaReference={selectFormulaReference} onResizeColumn={setColumnWidth} onResizeRow={setRowHeight} onSelect={selectCell} onSelectAll={selectAll} onSelectColumn={selectColumn} onSelectRow={selectRow} onViewport={onViewport} onVisibleRange={onVisibleRange} onWheelExtend={onWheelExtend} presence={presence} selection={selection} selections={extraSelections} sheet={sheet} zoom={zoom} />
      </section>
      <div className="sheet-tabs infinite-sheet-tabs"><button aria-label="新增工作表" className="sheet-add-button" onClick={() => void addSheet()} type="button"><Plus /></button>{sheets.map((item) => <button className={`sheet-tab${item.id === activeSheetId ? " active" : ""}`} key={item.id} onClick={() => setActiveSheetId(item.id)} onDoubleClick={() => { if (item.id === activeSheetId) void renameActiveSheet(); }} onContextMenu={(event) => { event.preventDefault(); if (item.id === activeSheetId) void renameActiveSheet(); }} type="button">{item.name}</button>)}<span className="sheet-tab-spacer" /><button aria-label="删除工作表" className="infinite-sheet-delete" onClick={() => void removeActiveSheet()} title="删除当前工作表" type="button"><X /></button></div>
      <div className="sheet-statistics infinite-sheet-statistics"><div className="sheet-statistics-values"><span>平均值={Number(statistics.average.toFixed(2))}</span><span>计数={statistics.count}</span><span>求和={statistics.sum}</span></div><div className="sheet-zoom-controls"><output>{zoom}%</output><button aria-label="缩小" onClick={() => setZoom((current) => Math.max(50, current - 10))} type="button"><Minus /></button><input aria-label="缩放" max={200} min={50} onChange={(event) => setZoom(Number(event.target.value))} step={10} type="range" value={zoom} /><button aria-label="放大" onClick={() => setZoom((current) => Math.min(200, current + 10))} type="button"><Plus /></button></div></div>
    </section></main>
    {contextMenu ? <div className="context-menu infinite-context-menu" onMouseDown={(event) => event.stopPropagation()} role="menu" style={{ left: Math.min(contextMenu.x, window.innerWidth - 210), top: Math.min(contextMenu.y, window.innerHeight - 340) }}><button onClick={() => void copySelection()} type="button"><Copy />复制</button><button onClick={() => void cutSelection()} type="button"><Scissors />剪切</button><button onClick={() => void pasteFromClipboard()} type="button"><ClipboardPaste />粘贴</button><button onClick={() => void clearSelection()} type="button"><Eraser />清空内容</button><hr /><button onClick={mergedInSelection.length ? unmergeSelection : mergeSelection} type="button">{mergedInSelection.length ? "取消合并" : "合并单元格"}</button><button onClick={() => insertOrDelete("insert_rows")} type="button">上方插入行</button><button onClick={() => insertOrDelete("delete_rows")} type="button">删除选中行</button><button onClick={() => insertOrDelete("insert_columns")} type="button">左侧插入列</button><button onClick={() => insertOrDelete("delete_columns")} type="button">删除选中列</button></div> : null}
    {toast ? <div className="toast" role="status">{toast}</div> : null}
  </div>;
}
