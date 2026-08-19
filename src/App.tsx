import {
  type CSSProperties,
  type ClipboardEvent,
  type FormEvent,
  type KeyboardEvent,
  type MouseEvent as ReactMouseEvent,
  type ReactNode,
  Fragment,
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
  ArrowDown,
  ArrowUp,
  ClipboardPaste,
  Copy,
  ChevronDown,
  Eraser,
  Grid3X3,
  Minus,
  PaintBucket,
  Paintbrush,
  Palette,
  Pencil,
  Plus,
  Redo2,
  Scissors,
  Search,
  Snowflake,
  Square,
  MoveHorizontal,
  MoveVertical,
  TableCellsMerge,
  TableCellsSplit,
  TableProperties,
  Trash2,
  Undo2,
  WrapText,
  X,
} from "lucide-react";
import {
  ApiError,
  addRow,
  deleteRows,
  ensureGrid,
  eventsUrl,
  formatCells,
  getState,
  submitOperations,
  updateCell,
  updateCells,
  updateLayout,
  updateBorders,
  type BorderAction,
  updatePresence,
} from "./api";
import {
  parseHtmlSpreadsheetClipboard,
  parseSpreadsheetClipboard,
  serializeSpreadsheetClipboard,
} from "./clipboard";
import type {
  ActivityItem,
  CellBorders,
  CellEvent,
  CellFormat,
  CellMeta,
  CollaborationIdentity,
  EditableField,
  LayoutEvent,
  PresenceItem,
  RowEvent,
  GridBookRow,
  GridBookState,
} from "./types";
import { displayFormulaValue } from "./formulas";
import { buildFillPreview, fillCellAt, type FillBounds, type FillMode } from "./fill";
import InfiniteWorkbook from "./InfiniteWorkbook";

const CLIENT_KEY = "gridbook-client-id";
const NAME_KEY = "gridbook-display-name";

const FONT_OPTIONS = [
  { value: "Arial", label: "Arial" },
  { value: "Microsoft YaHei", label: "微软雅黑" },
  { value: "SimSun", label: "宋体" },
  { value: "FangSong", label: "仿宋" },
  { value: "仿宋_GB2312", label: "仿宋_GB2312" },
  { value: "KaiTi", label: "楷体" },
  { value: "楷体_GB2312", label: "楷体_GB2312" },
  { value: "SimHei", label: "黑体" },
  { value: "Times New Roman", label: "Times New Roman" },
] as const;

const RECENT_COLORS = ["#ffffff", "#d9eaf7", "#fff2cc", "#e2f0d9", "#f4cccc", "#d9d9d9", "#c6e0b4", "#ffd966", "#5b9bd5", "#c00000"] as const;
const PALETTE_THEME_ROWS = [
  ["#ffffff", "#000000", "#1f4e78", "#4472c4", "#ed7d31", "#a5a5a5", "#ffc000", "#70ad47", "#5b9bd5", "#c00000"],
  ["#f2f2f2", "#7f7f7f", "#d9eaf7", "#d9e2f3", "#fce4d6", "#e7e6e6", "#fff2cc", "#e2f0d9", "#ddebf7", "#f4cccc"],
  ["#d9d9d9", "#595959", "#bdd7ee", "#b4c7e7", "#f8cbad", "#d0cece", "#ffe699", "#c6e0b4", "#bdd7ee", "#ea9999"],
  ["#bfbfbf", "#404040", "#9dc3e6", "#8eaadb", "#f4b183", "#afabab", "#ffd966", "#a9d18e", "#9dc3e6", "#e06666"],
  ["#7f7f7f", "#262626", "#5b9bd5", "#4472c4", "#ed7d31", "#7f7f7f", "#ffc000", "#70ad47", "#5b9bd5", "#c00000"],
] as const;

interface FieldDefinition {
  key: EditableField;
  label: string;
  shortLabel?: string;
  type: "text" | "number";
  width: number;
}

function excelColumnName(index: number): string {
  let value = index + 1;
  let label = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

interface CellPosition {
  row: number;
  col: number;
}

interface FormulaEditorHandle {
  beginReference: (cell: CellPosition) => void;
  updateReference: (start: CellPosition, end: CellPosition) => void;
}

interface FormulaReferenceToken {
  text: string;
  start: number;
  end: number;
  startCell: CellPosition;
  endCell: CellPosition;
  color: string;
  fill: string;
}

interface CellSelection {
  anchor: CellPosition;
  focus: CellPosition;
}

type EditableCellValue = string | number | null;

interface ValueHistoryItem {
  updates: Array<{
    row_id: number;
    field: EditableField;
    before: EditableCellValue;
    after: EditableCellValue;
  }>;
  label: string;
}

interface FormatHistoryItem {
  cells: Array<{ row_id: number; field: EditableField }>;
  before: CellFormat[];
  after: CellFormat;
}

type HistoryAction =
  | { kind: "value"; item: ValueHistoryItem }
  | { kind: "format"; item: FormatHistoryItem };

function formulaReference(start: CellPosition, end = start): string {
  const first = `${excelColumnName(start.col)}${start.row + 1}`;
  const last = `${excelColumnName(end.col)}${end.row + 1}`;
  return first === last ? first : `${first}:${last}`;
}

function formulaCursorRange(input: HTMLInputElement | null, value: string): { start: number; end: number } {
  const start = input?.selectionStart ?? value.length;
  const end = input?.selectionEnd ?? start;
  if (start === end && start === value.length && value.trimStart().startsWith("=")) {
    const trimmed = value.trimEnd();
    if (trimmed.endsWith(")")) {
      const closing = value.lastIndexOf(")", trimmed.length - 1);
      if (closing >= 0) return { start: closing, end: closing };
    }
  }
  return { start, end };
}

const FORMULA_REFERENCE_COLORS = [
  { color: "#2f6fed", fill: "rgba(47, 111, 237, .14)" },
  { color: "#d84a4a", fill: "rgba(216, 74, 74, .14)" },
  { color: "#8a55c7", fill: "rgba(138, 85, 199, .14)" },
  { color: "#159570", fill: "rgba(21, 149, 112, .14)" },
  { color: "#d9822b", fill: "rgba(217, 130, 43, .15)" },
  { color: "#c4478f", fill: "rgba(196, 71, 143, .14)" },
] as const;

function excelColumnIndex(label: string): number {
  let value = 0;
  for (const character of label.toUpperCase()) value = value * 26 + character.charCodeAt(0) - 64;
  return value - 1;
}

function parseFormulaCellAddress(value: string): CellPosition | null {
  const match = value.match(/^\$?([A-Z]{1,3})\$?(\d+)$/i);
  if (!match) return null;
  const row = Number(match[2]) - 1;
  const col = excelColumnIndex(match[1]);
  return Number.isFinite(row) && row >= 0 && col >= 0 ? { row, col } : null;
}

function parseFormulaReferences(value: string): FormulaReferenceToken[] {
  if (!value.trimStart().startsWith("=")) return [];
  const tokens: FormulaReferenceToken[] = [];
  const pattern = /(^|[^A-Za-z0-9_])(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)/gi;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(value))) {
    const text = match[2];
    const parts = text.split(":");
    const startCell = parseFormulaCellAddress(parts[0]);
    const endCell = parseFormulaCellAddress(parts[1] ?? parts[0]);
    if (!startCell || !endCell) continue;
    const color = FORMULA_REFERENCE_COLORS[tokens.length % FORMULA_REFERENCE_COLORS.length];
    tokens.push({
      text,
      start: match.index + match[1].length,
      end: match.index + match[1].length + text.length,
      startCell: { row: Math.min(startCell.row, endCell.row), col: Math.min(startCell.col, endCell.col) },
      endCell: { row: Math.max(startCell.row, endCell.row), col: Math.max(startCell.col, endCell.col) },
      color: color.color,
      fill: color.fill,
    });
  }
  return tokens;
}

function FormulaReferencePreview({ value }: { value: string }) {
  const references = parseFormulaReferences(value);
  if (!value) return null;
  if (!references.length) return <span>{value}</span>;
  const nodes: ReactNode[] = [];
  let cursor = 0;
  references.forEach((reference, index) => {
    if (reference.start > cursor) nodes.push(<span key={`formula-text-${index}`}>{value.slice(cursor, reference.start)}</span>);
    nodes.push(<span key={`formula-reference-${index}`} style={{ color: reference.color }}>{value.slice(reference.start, reference.end)}</span>);
    cursor = reference.end;
  });
  if (cursor < value.length) nodes.push(<span key="formula-text-tail">{value.slice(cursor)}</span>);
  return <>{nodes}</>;
}

const FORMULA_FUNCTIONS = [
  "SUM", "AVERAGE", "COUNT", "COUNTA", "MAX", "MIN", "IF", "AND", "OR", "NOT", "ROUND", "ROUNDUP", "ROUNDDOWN",
] as const;

function formulaSuggestionQuery(value: string): string | null {
  if (!value.trimStart().startsWith("=")) return null;
  const match = value.match(/(?:^|[=+\-*/(,;])([A-Za-z]*)$/);
  return match ? match[1].toUpperCase() : null;
}

function RoundedTextFormatIcon({ kind }: { kind: "bold" | "italic" | "underline" }) {
  return (
    <svg aria-hidden="true" fill="none" height="16" viewBox="0 0 24 24" width="16" xmlns="http://www.w3.org/2000/svg" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="2">
      {kind === "bold" ? <path d="M7 5h5a3.5 3.5 0 0 1 0 7H7m0-7v14h6a3.5 3.5 0 0 0 0-7H7" /> : null}
      {kind === "italic" ? <><path d="M10 5h8" /><path d="M6 19h8" /><path d="M15 5 9 19" /></> : null}
      {kind === "underline" ? <><path d="M7 5v6a5 5 0 0 0 10 0V5" /><path d="M5 19h14" /></> : null}
    </svg>
  );
}

function FormulaSuggestions({ value, onChoose }: { value: string; onChoose: (name: string) => void }) {
  const query = formulaSuggestionQuery(value);
  if (query === null) return null;
  const suggestions = FORMULA_FUNCTIONS.filter((name) => name.startsWith(query));
  if (!suggestions.length) return null;
  return (
    <div
      className="formula-suggestions"
      role="listbox"
      aria-label="Formula functions"
      onWheel={(event) => {
        // Keep formula-list scrolling independent from the spreadsheet viewport.
        // Without stopping propagation, wheel events bubble to the virtual grid
        // and move the sheet while the user is browsing functions.
        event.stopPropagation();
        const menu = event.currentTarget;
        const atTop = menu.scrollTop <= 0;
        const atBottom = menu.scrollTop + menu.clientHeight >= menu.scrollHeight - 1;
        const scrollingUp = event.deltaY < 0;
        if ((scrollingUp && atTop) || (!scrollingUp && atBottom)) {
          // Prevent scroll chaining when the list reaches either edge.
          event.preventDefault();
        }
      }}
    >
      {suggestions.map((name) => (
        <button key={name} onMouseDown={(event) => { event.preventDefault(); onChoose(name); }} role="option" type="button">
          <span className="suggestion-fx" aria-hidden="true">fx</span>{name}
        </button>
      ))}
    </div>
  );
}

interface ContextMenuState extends CellPosition {
  x: number;
  y: number;
}

interface SheetContextMenuState {
  sheet: string;
  x: number;
  y: number;
}

const FIELDS: FieldDefinition[] = [
  { key: "account_name", label: "A", type: "text", width: 176 },
  { key: "teacher", label: "B", type: "text", width: 96 },
  { key: "channel", label: "C", type: "text", width: 96 },
  { key: "publish_count", label: "D", type: "number", width: 104 },
  { key: "total_followers", label: "E", type: "number", width: 116 },
  {
    key: "short_graphic_leads",
    label: "F",
    type: "number",
    width: 154,
  },
  { key: "live_leads", label: "G", type: "number", width: 112 },
  { key: "openings", label: "H", type: "number", width: 94 },
  { key: "remark", label: "I", type: "text", width: 220 },
  ...Array.from({ length: 91 }, (_, index) => ({
    key: `custom_${index + 1}` as EditableField,
    label: excelColumnName(9 + index),
    type: "text" as const,
    width: 132,
  })),
];

const INITIAL_DATA_FIELD_COUNT = 9;
// Keep only a small editable tail in the initial viewport. More rows are
// appended when the user approaches the bottom, so the scrollbar grows with
// actual exploration instead of starting with hundreds of blank rows.
const GRID_ROW_TARGET = 200;
const GRID_ROW_MAX = 10000;
// Keep the first request small even when another collaborator has already
// expanded the sheet to thousands of rows.  More rows are fetched in chunks
// as the user scrolls, so entering the document never waits for the full
// persisted workbook.
const INITIAL_STATE_LIMIT = GRID_ROW_TARGET;
const PRIMARY_SHEET_NAME = "工作表1";
const DEFAULT_SHEET_NAMES = [PRIMARY_SHEET_NAME, "工作表2", "工作表3", "工作表4"] as const;
const SHEET_NAMES_STORAGE_KEY = "gridbook-sheet-names";
// New worksheets use a neutral spreadsheet grid instead of inheriting the
// data-specific widths from the primary contribution sheet.
const DEFAULT_WORKSHEET_COLUMN_WIDTH = 120;
const DEFAULT_WORKSHEET_ROW_HEIGHT = 29;
const LEGACY_DEFAULT_ROW_HEIGHT = 55;

const EMPTY_STATE: GridBookState = {
  schema_version: 1,
  revision: 0,
  next_row_id: 1,
  period: { start: "", end: "" },
  rows: [],
  cells: {},
  layout: { column_widths: {}, row_heights: {}, merges: [], formats: {} },
  activity: [],
  presence: [],
};

function createBlankRow(id: number): GridBookRow {
  return {
    id,
    account_name: "",
    teacher: "",
    channel: "",
    publish_count: null,
    total_followers: null,
    short_graphic_leads: null,
    live_leads: null,
    openings: null,
    remark: "",
  };
}

function createBlankSheetState(rowCount = GRID_ROW_TARGET): GridBookState {
  const rows = Array.from({ length: rowCount }, (_, index) => createBlankRow(index + 1));
  const columnWidths = FIELDS.reduce<Partial<Record<EditableField, number>>>((result, field) => {
    result[field.key] = DEFAULT_WORKSHEET_COLUMN_WIDTH;
    return result;
  }, {});
  return {
    ...EMPTY_STATE,
    rows,
    next_row_id: rowCount + 1,
    layout: { column_widths: columnWidths, row_heights: {}, merges: [], formats: {} },
  };
}

const AVATAR_COLORS = [
  "#4169e1",
  "#8b5cf6",
  "#0e8f72",
  "#d05b38",
  "#b17814",
  "#c44670",
  "#367da6",
];

function createClientId(): string {
  const existing = localStorage.getItem(CLIENT_KEY);
  if (existing) return existing;
  const suffix =
    typeof crypto.randomUUID === "function"
      ? crypto.randomUUID()
      : `${Date.now()}-${Math.random().toString(16).slice(2)}`;
  const clientId = `client-${suffix}`;
  localStorage.setItem(CLIENT_KEY, clientId);
  return clientId;
}

function colorFor(seed: string): string {
  let hash = 0;
  for (let index = 0; index < seed.length; index += 1) {
    hash = (hash * 31 + seed.charCodeAt(index)) | 0;
  }
  return AVATAR_COLORS[Math.abs(hash) % AVATAR_COLORS.length];
}

function avatarStyle(seed: string): CSSProperties {
  return { "--avatar-color": colorFor(seed) } as CSSProperties;
}

function initials(name: string): string {
  const clean = name.trim();
  if (!clean) return "协";
  return clean.length <= 2 ? clean : clean.slice(-2);
}

function formatValue(value: unknown, type: FieldDefinition["type"], format?: CellFormat): string {
  if (value === null || value === undefined || value === "") return "";
  const raw = String(value);
  if (format?.number_format === "text") return raw;
  if (format?.number_format === "date") {
    const parsed = new Date(raw.includes("T") ? raw : `${raw}T00:00:00`);
    if (!Number.isNaN(parsed.getTime())) {
      return `${parsed.getFullYear()}/${parsed.getMonth() + 1}/${parsed.getDate()}`;
    }
    return raw;
  }
  const numeric = Number(raw.replaceAll(",", "").replace(/%$/, ""));
  if (format?.number_format === "percentage") {
    if (!Number.isFinite(numeric)) return raw;
    return new Intl.NumberFormat("zh-CN", { maximumFractionDigits: 2 }).format(raw.endsWith("%") ? numeric : numeric * 100) + "%";
  }
  if (format?.number_format === "number") {
    return Number.isFinite(numeric)
      ? new Intl.NumberFormat("zh-CN", { maximumFractionDigits: 2 }).format(numeric)
      : raw;
  }
  if (type === "number" && typeof value === "number") {
    return new Intl.NumberFormat("zh-CN").format(value);
  }
  return raw;
}

function formatPeriod(start: string, end: string): string {
  const options: Intl.DateTimeFormatOptions = { month: "long", day: "numeric" };
  const formatter = new Intl.DateTimeFormat("zh-CN", options);
  return `${formatter.format(new Date(`${start}T00:00:00+08:00`))} – ${formatter.format(
    new Date(`${end}T00:00:00+08:00`),
  )}`;
}

function timeLabel(value: string): string {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "刚刚";
  return new Intl.DateTimeFormat("zh-CN", {
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  }).format(date);
}

function fullTime(value: string): string {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  return new Intl.DateTimeFormat("zh-CN", {
    month: "numeric",
    day: "numeric",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  }).format(date);
}

function selectionBounds(selection: CellSelection) {
  return {
    top: Math.min(selection.anchor.row, selection.focus.row),
    bottom: Math.max(selection.anchor.row, selection.focus.row),
    left: Math.min(selection.anchor.col, selection.focus.col),
    right: Math.max(selection.anchor.col, selection.focus.col),
  };
}

function selectionLabel(selections: CellSelection[]): string {
  return selections.map((selection) => {
    const bounds = selectionBounds(selection);
    const start = `${excelColumnName(bounds.left)}${bounds.top + 1}`;
    const end = `${excelColumnName(bounds.right)}${bounds.bottom + 1}`;
    return start === end ? start : `${start}:${end}`;
  }).join(", ");
}

function Icon({ name }: { name: "plus" | "download" | "users" | "edit" | "check" }) {
  if (name === "plus") {
    return <svg viewBox="0 0 24 24"><path d="M12 5v14M5 12h14" /></svg>;
  }
  if (name === "download") {
    return (
      <svg viewBox="0 0 24 24">
        <path d="M12 3v12m0 0 4-4m-4 4-4-4M5 19h14" />
      </svg>
    );
  }
  if (name === "users") {
    return (
      <svg viewBox="0 0 24 24">
        <path d="M16 18.5c0-2-1.8-3.5-4-3.5s-4 1.5-4 3.5M12 12a3 3 0 1 0 0-6 3 3 0 0 0 0 6ZM18 12.5c1.8.2 3 1.3 3 2.8M17 5.7a2.5 2.5 0 0 1 0 4.8M6 12.5c-1.8.2-3 1.3-3 2.8M7 5.7a2.5 2.5 0 0 0 0 4.8" />
      </svg>
    );
  }
  if (name === "check") {
    return <svg viewBox="0 0 24 24"><path d="m5 12.5 4.2 4.2L19 7" /></svg>;
  }
  return (
    <svg viewBox="0 0 24 24">
      <path d="m14.5 5.5 4 4M5 19l3.8-.8L19 8a1.4 1.4 0 0 0 0-2l-1-1a1.4 1.4 0 0 0-2 0L5.8 15.2 5 19Z" />
    </svg>
  );
}

function Avatar({
  clientId,
  name,
  size = "normal",
}: {
  clientId: string;
  name: string;
  size?: "small" | "normal" | "large";
}) {
  return (
    <span
      className={`avatar avatar-${size}`}
      style={avatarStyle(clientId)}
      title={name}
    >
      {initials(name)}
    </span>
  );
}

function IdentityDialog({
  currentName,
  onSave,
  canClose,
  onClose,
}: {
  currentName: string;
  onSave: (name: string) => void;
  canClose: boolean;
  onClose: () => void;
}) {
  const [name, setName] = useState(currentName);
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    setName(currentName);
    const timer = window.setTimeout(() => inputRef.current?.focus(), 80);
    return () => window.clearTimeout(timer);
  }, [currentName]);

  const submit = (event: FormEvent) => {
    event.preventDefault();
    const normalized = name.trim();
    if (normalized) onSave(normalized.slice(0, 24));
  };

  return (
    <div className="dialog-backdrop" role="presentation" onMouseDown={canClose ? onClose : undefined}>
      <form
        aria-labelledby="identity-title"
        className="identity-dialog"
        onMouseDown={(event) => event.stopPropagation()}
        onSubmit={submit}
        role="dialog"
      >
        <div className="dialog-mark" aria-hidden="true">
          <span /><span /><span /><span />
        </div>
        <p className="dialog-eyebrow">GRIDBOOK COLLABORATION</p>
        <h2 id="identity-title">先告诉大家你是谁</h2>
        <p className="dialog-copy">你的名字会显示在正在编辑的单元格和协作动态中。</p>
        <label className="name-field">
          <span>显示名称</span>
          <input
            autoComplete="name"
            maxLength={24}
            onChange={(event) => setName(event.target.value)}
            placeholder="例如：王语桐"
            ref={inputRef}
            value={name}
          />
        </label>
        <button className="button primary dialog-submit" disabled={!name.trim()} type="submit">
          进入协作表
        </button>
            <p className="dialog-note">独立协作表格环境</p>
      </form>
    </div>
  );
}

interface EditableCellProps {
  field: FieldDefinition;
  identity: CollaborationIdentity;
  meta?: CellMeta;
  onCommit: (value: string | number | null, version: number) => Promise<void>;
  onFocusCell: (cell: { row_id: number; field: EditableField } | null) => void;
  onMoveAfterCommit: (rowDelta: number, columnDelta: number) => void;
  onOpenContextMenu: (event: ReactMouseEvent) => void;
  onSelect: (extend: boolean, additive: boolean) => void;
  onSelectionDrag: () => void;
  onFormulaEditorChange?: (handle: FormulaEditorHandle | null) => void;
  onFormulaDraftChange?: (value: string) => void;
  onFormulaReferencePointerDown?: () => boolean;
  onFormulaReferencePointerMove?: () => boolean;
  remoteEditor?: PresenceItem;
  remoteEditorInHeader?: boolean;
  row: GridBookRow;
  flash: boolean;
  selected: boolean;
  active: boolean;
  selectionEdges: { top: boolean; right: boolean; bottom: boolean; left: boolean };
  editRequest?: { id: number; seed: string | null };
  format?: CellFormat;
  displayValue?: unknown;
}

function EditableCell({
  field,
  identity,
  meta,
  onCommit,
  onFocusCell,
  onMoveAfterCommit,
  onOpenContextMenu,
  onSelect,
  onSelectionDrag,
  onFormulaEditorChange = () => undefined,
  onFormulaDraftChange = () => undefined,
  onFormulaReferencePointerDown = () => false,
  onFormulaReferencePointerMove = () => false,
  remoteEditor,
  remoteEditorInHeader = false,
  row,
  flash,
  selected,
  active,
  selectionEdges,
  editRequest,
  format,
  displayValue,
}: EditableCellProps) {
  const value = row[field.key];
  const [editing, setEditing] = useState(false);
  const [saving, setSaving] = useState(false);
  const [draft, setDraft] = useState(value === null || value === undefined ? "" : String(value));
  const cancelled = useRef(false);
  const moveAfterCommit = useRef<[number, number] | null>(null);
  const directSeed = useRef<string | null>(null);
  const lastEditRequest = useRef(0);
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (!editing) setDraft(value === null || value === undefined ? "" : String(value));
  }, [editing, value]);

  useEffect(() => {
    if (!editing) return;
    const input = inputRef.current;
    if (!input) return;
    input.focus();
    if (directSeed.current === null) {
      if (input.value.trimStart().startsWith("=") && input.value.trimEnd().endsWith(")")) {
        const closing = input.value.lastIndexOf(")", input.value.trimEnd().length - 1);
        input.setSelectionRange(closing >= 0 ? closing : input.value.length, closing >= 0 ? closing : input.value.length);
      } else {
        input.select();
      }
    } else input.setSelectionRange(input.value.length, input.value.length);
    directSeed.current = null;
  }, [editing]);

  const startEditing = (seed: string | null = null) => {
    if (remoteEditor) return;
    cancelled.current = false;
    directSeed.current = seed;
    if (seed !== null) setDraft(seed);
    setEditing(true);
    onFocusCell({ row_id: row.id, field: field.key });
  };

  const chooseFormula = (name: string) => {
    const input = inputRef.current;
    const cursor = input?.selectionStart ?? draft.length;
    const end = input?.selectionEnd ?? cursor;
    const before = draft.slice(0, cursor);
    const token = before.match(/[A-Za-z]*$/)?.[0] ?? "";
    const tokenStart = before.length - token.length;
    const next = `${before.slice(0, tokenStart)}${name}()${draft.slice(end)}`;
    setDraft(next);
    requestAnimationFrame(() => {
      inputRef.current?.focus();
      const nextCursor = tokenStart + name.length + 1;
      inputRef.current?.setSelectionRange(nextCursor, nextCursor);
    });
  };

  useEffect(() => {
    if (!editRequest || editRequest.id === lastEditRequest.current) return;
    lastEditRequest.current = editRequest.id;
    startEditing(editRequest.seed);
    // startEditing intentionally uses the current row/cell closure.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [editRequest]);

  useEffect(() => {
    if (!editing) return;
    if (!draft.trimStart().startsWith("=")) {
      onFormulaEditorChange(null);
      return;
    }
    const insertReference = (start: number, end: number, cell: CellPosition) => {
      // The caller has already normalized the active range (including the
      // special case that places a new reference immediately before the
      // formula's closing parenthesis).  Re-reading selectionStart here can
      // race the browser's caret update and shift the insertion point.
      const prefix = draft.slice(0, start);
      const suffix = draft.slice(end);
      const reference = `${excelColumnName(cell.col)}${cell.row + 1}`;
      setDraft(`${prefix}${reference}${suffix}`);
      requestAnimationFrame(() => {
        inputRef.current?.focus();
        const nextCursor = prefix.length + reference.length;
        inputRef.current?.setSelectionRange(nextCursor, nextCursor);
      });
      return { prefix, suffix, cursorStart: start, cursorEnd: end };
    };
    let referenceBase: { prefix: string; suffix: string } | null = null;
    const replaceReference = (start: CellPosition, end: CellPosition) => {
      if (!referenceBase) return;
      const reference = formulaReference(start, end);
      setDraft(`${referenceBase.prefix}${reference}${referenceBase.suffix}`);
      requestAnimationFrame(() => {
        inputRef.current?.focus();
        const nextCursor = referenceBase!.prefix.length + reference.length;
        inputRef.current?.setSelectionRange(nextCursor, nextCursor);
      });
    };
    const handle: FormulaEditorHandle = {
      beginReference: (cell) => {
        const input = inputRef.current;
        const cursorRange = formulaCursorRange(input, draft);
        const cursorStart = cursorRange.start;
        const cursorEnd = cursorRange.end;
        const prefix = draft.slice(0, cursorStart);
        const suffix = draft.slice(cursorEnd);
        referenceBase = { prefix, suffix };
        insertReference(cursorStart, cursorEnd, cell);
      },
      updateReference: replaceReference,
    };
    onFormulaEditorChange(handle);
    return () => onFormulaEditorChange(null);
  }, [draft, editing, onFormulaEditorChange]);

  useEffect(() => {
    onFormulaDraftChange(editing ? draft : "");
    return () => onFormulaDraftChange("");
  }, [draft, editing, onFormulaDraftChange]);

  const commit = async () => {
    if (cancelled.current) {
      cancelled.current = false;
      setEditing(false);
      onFocusCell(null);
      return;
    }
    const nextValue = draft;
    const currentValue = value === undefined ? null : value;
    if (nextValue === currentValue) {
      setEditing(false);
      onFocusCell(null);
      const move = moveAfterCommit.current;
      moveAfterCommit.current = null;
      if (move) onMoveAfterCommit(move[0], move[1]);
      return;
    }
    setSaving(true);
    try {
      await onCommit(nextValue, meta?.version ?? 0);
      setEditing(false);
      onFocusCell(null);
      const move = moveAfterCommit.current;
      moveAfterCommit.current = null;
      if (move) onMoveAfterCommit(move[0], move[1]);
    } catch {
      inputRef.current?.focus();
    } finally {
      setSaving(false);
    }
  };

  const onKeyDown = (event: KeyboardEvent<HTMLInputElement>) => {
    if (event.key === "Enter") {
      event.preventDefault();
      moveAfterCommit.current = [event.shiftKey ? -1 : 1, 0];
      event.currentTarget.blur();
    }
    if (event.key === "Tab") {
      event.preventDefault();
      moveAfterCommit.current = [0, event.shiftKey ? -1 : 1];
      event.currentTarget.blur();
    }
    if (event.key === "Escape") {
      event.preventDefault();
      cancelled.current = true;
      moveAfterCommit.current = null;
      setDraft(value === null || value === undefined ? "" : String(value));
      event.currentTarget.blur();
    }
  };

  const classNames = [
    "editable-cell",
    editing ? "is-editing" : "",
    remoteEditor ? "has-remote-editor" : "",
    flash ? "is-flashing" : "",
    saving ? "is-saving" : "",
    selected ? "is-selected" : "",
    active ? "is-active-cell" : "",
    format?.fill && format.fill !== "none" ? "has-cell-fill" : "",
    selectionEdges.top ? "selection-top" : "",
    selectionEdges.right ? "selection-right" : "",
    selectionEdges.bottom ? "selection-bottom" : "",
    selectionEdges.left ? "selection-left" : "",
  ]
    .filter(Boolean)
    .join(" ");
  const title = meta
    ? `由 ${meta.updated_by} 编辑于 ${fullTime(meta.updated_at)}`
    : "点击编辑";

  return (
    <div
      className={classNames}
      aria-label={meta ? title : undefined}
      style={{
        backgroundColor: format?.fill && format.fill !== "none" ? format.fill : undefined,
        ["--cell-fill" as string]: format?.fill && format.fill !== "none" ? format.fill : "transparent",
        ...(remoteEditor ? avatarStyle(remoteEditor.client_id) : {}),
      } as CSSProperties}
    >
      {editing ? (
        <div className="cell-formula-wrap">
          {draft.trimStart().startsWith("=") ? <span className="formula-reference-preview" aria-hidden="true"><FormulaReferencePreview value={draft} /></span> : null}
          <input
          className={`cell-editor-input formula-editor-input${draft.trimStart().startsWith("=") ? " has-formula-references" : ""}`}
          aria-label={`编辑${field.label}`}
          disabled={saving}
          onBlur={() => void commit()}
          onChange={(event) => setDraft(event.target.value)}
          onKeyDown={onKeyDown}
          ref={inputRef}
          type="text"
          value={draft}
          />
        </div>
      ) : (
        <button
          aria-label={`选择${field.label}；双击编辑`}
          onContextMenu={onOpenContextMenu}
          onDoubleClick={() => startEditing(null)}
          onMouseDown={(event) => {
            if (event.button !== 0) return;
            event.preventDefault();
            if (onFormulaReferencePointerDown()) return;
            onSelect(event.shiftKey, event.ctrlKey || event.metaKey);
          }}
          onMouseEnter={() => {
            if (!onFormulaReferencePointerMove()) onSelectionDrag();
          }}
          onKeyDown={(event) => {
            if (event.key === "Enter" || event.key === "F2") {
              event.preventDefault();
              startEditing();
            }
          }}
          style={{
            fontSize: format?.font_size ? `calc(${format.font_size}px * var(--sheet-zoom, 1))` : undefined,
            fontFamily: format?.font_family || undefined,
            fontWeight: format?.bold ? 700 : undefined,
            fontStyle: format?.italic === true ? "italic" : "normal",
            textDecoration: format?.underline ? "underline" : undefined,
            color: format?.font_color && format.font_color !== "default" ? format.font_color : undefined,
          }}
          type="button"
        >
          <span className={value === null || value === undefined || value === "" ? "empty-value" : ""}>
            {formatValue(displayValue === undefined ? value : displayValue, field.type, format)}
          </span>
        </button>
      )}
      {editing ? <FormulaSuggestions value={draft} onChoose={chooseFormula} /> : null}
      {remoteEditor && !remoteEditorInHeader ? (
        <span
          aria-label={`${remoteEditor.name} 正在输入`}
          className="remote-editor"
          style={avatarStyle(remoteEditor.client_id)}
        >
          {remoteEditor.name} 正在输入
        </span>
      ) : null}
      {saving ? <span className="saving-indicator" aria-label="正在保存" /> : null}
    </div>
  );
}

function SpreadsheetContextMenu({
  menu,
  selection,
  meta,
  onClose,
  onCopy,
  onCut,
  onPaste,
  onClear,
  onMerge,
  onUnmerge,
  onInsertAbove,
  onInsertBelow,
  onDeleteRows,
  onResetRowHeight,
  onResetColumnWidth,
}: {
  menu: ContextMenuState;
  selection: CellSelection;
  meta?: CellMeta;
  onClose: () => void;
  onCopy: () => void;
  onCut: () => void;
  onPaste: () => void;
  onClear: () => void;
  onMerge: () => void;
  onUnmerge: () => void;
  onInsertAbove: () => void;
  onInsertBelow: () => void;
  onDeleteRows: () => void;
  onResetRowHeight: () => void;
  onResetColumnWidth: () => void;
}) {
  const menuRef = useRef<HTMLDivElement>(null);
  const [position, setPosition] = useState({ left: menu.x, top: menu.y });

  const clampPosition = useCallback(() => {
    const node = menuRef.current;
    if (!node) return;
    const margin = 10;
    const rect = node.getBoundingClientRect();
    const maxLeft = Math.max(margin, window.innerWidth - rect.width - margin);
    const maxTop = Math.max(margin, window.innerHeight - rect.height - margin);
    setPosition({
      left: Math.min(Math.max(margin, menu.x), maxLeft),
      top: Math.min(Math.max(margin, menu.y), maxTop),
    });
  }, [menu.x, menu.y]);

  useLayoutEffect(() => {
    clampPosition();
    window.addEventListener("resize", clampPosition);
    return () => window.removeEventListener("resize", clampPosition);
  }, [clampPosition]);

  useEffect(() => {
    const close = (event: globalThis.KeyboardEvent) => {
      if (event.key === "Escape") onClose();
    };
    window.addEventListener("keydown", close);
    return () => {
      window.removeEventListener("keydown", close);
    };
  }, [onClose]);

  return (
    <div
      className="context-menu"
      onContextMenu={(event) => event.preventDefault()}
      onMouseDown={(event) => event.stopPropagation()}
      onWheel={(event) => event.stopPropagation()}
      ref={menuRef}
      role="menu"
      style={{
        left: position.left,
        top: position.top,
      }}
    >
      <div className="context-menu-group-label">编辑</div>
      <button onClick={onCopy} role="menuitem" type="button">
        <span className="menu-icon"><Copy aria-hidden="true" /></span>
        <span>复制</span>
        <kbd>Ctrl C</kbd>
      </button>
      <button onClick={onCut} role="menuitem" type="button">
        <span className="menu-icon"><Scissors aria-hidden="true" /></span>
        <span>剪切</span>
        <kbd>Ctrl X</kbd>
      </button>
      <button onClick={onPaste} role="menuitem" type="button">
        <span className="menu-icon"><ClipboardPaste aria-hidden="true" /></span>
        <span>粘贴</span>
        <kbd>Ctrl V</kbd>
      </button>
      <div className="context-menu-separator" />
      <div className="context-menu-group-label">单元格</div>
      <button onClick={onMerge} role="menuitem" type="button">
        <span className="menu-icon"><TableCellsMerge aria-hidden="true" /></span>
        <span>合并单元格</span>
        <kbd>合并</kbd>
      </button>
      <button onClick={onUnmerge} role="menuitem" type="button">
        <span className="menu-icon"><TableCellsSplit aria-hidden="true" /></span>
        <span>取消合并</span>
        <kbd>拆分</kbd>
      </button>
      <div className="context-menu-separator" />
      <div className="context-menu-group-label">行列</div>
      <button onClick={onInsertAbove} role="menuitem" type="button">
        <span className="menu-icon"><ArrowUp aria-hidden="true" /></span>
        <span>上方插入行</span>
        <kbd>行</kbd>
      </button>
      <button onClick={onInsertBelow} role="menuitem" type="button">
        <span className="menu-icon"><ArrowDown aria-hidden="true" /></span>
        <span>下方插入行</span>
        <kbd>行</kbd>
      </button>
      <button onClick={onDeleteRows} role="menuitem" type="button">
        <span className="menu-icon"><Trash2 aria-hidden="true" /></span>
        <span>删除选中行</span>
        <kbd>行</kbd>
      </button>
      <button onClick={onResetRowHeight} role="menuitem" type="button">
        <span className="menu-icon"><MoveVertical aria-hidden="true" /></span>
        <span>恢复标准行高</span>
        <kbd>29</kbd>
      </button>
      <button onClick={onResetColumnWidth} role="menuitem" type="button">
        <span className="menu-icon"><MoveHorizontal aria-hidden="true" /></span>
        <span>恢复标准列宽</span>
        <kbd>列宽</kbd>
      </button>
      <div className="context-menu-separator" />
      <div className="context-menu-group-label">清除</div>
      <button className="clear-action" onClick={onClear} role="menuitem" type="button">
        <span className="menu-icon"><Eraser aria-hidden="true" /></span>
        <span>清空内容</span>
        <kbd>Delete</kbd>
      </button>
      {meta ? (
        <div className="context-menu-meta">
          最近由 <strong>{meta.updated_by}</strong> 更新 · {fullTime(meta.updated_at)}
        </div>
      ) : null}
    </div>
  );
}

function SheetContextMenu({
  menu,
  onClose,
  onRename,
  onDelete,
}: {
  menu: SheetContextMenuState;
  onClose: () => void;
  onRename: (sheetName: string) => void;
  onDelete: (sheetName: string) => void;
}) {
  const menuRef = useRef<HTMLDivElement>(null);
  const [position, setPosition] = useState({ left: menu.x, top: menu.y });

  const clampPosition = useCallback(() => {
    const node = menuRef.current;
    if (!node) return;
    const margin = 10;
    const rect = node.getBoundingClientRect();
    const maxLeft = Math.max(margin, window.innerWidth - rect.width - margin);
    const maxTop = Math.max(margin, window.innerHeight - rect.height - margin);
    setPosition({
      left: Math.min(Math.max(margin, menu.x), maxLeft),
      top: Math.min(Math.max(margin, menu.y), maxTop),
    });
  }, [menu.x, menu.y]);

  useLayoutEffect(() => {
    clampPosition();
    window.addEventListener("resize", clampPosition);
    return () => window.removeEventListener("resize", clampPosition);
  }, [clampPosition]);

  useEffect(() => {
    const close = (event: globalThis.KeyboardEvent) => {
      if (event.key === "Escape") onClose();
    };
    window.addEventListener("keydown", close);
    return () => window.removeEventListener("keydown", close);
  }, [onClose]);

  return (
    <div
      className="context-menu sheet-context-menu"
      onContextMenu={(event) => event.preventDefault()}
      onMouseDown={(event) => event.stopPropagation()}
      onWheel={(event) => event.stopPropagation()}
      ref={menuRef}
      role="menu"
      style={{ left: position.left, top: position.top }}
    >
      <div className="context-menu-title"><strong>{menu.sheet}</strong></div>
      <button onClick={() => onRename(menu.sheet)} role="menuitem" type="button">
        <span className="menu-icon"><Pencil aria-hidden="true" /></span>
        <span>重命名</span>
        <span />
      </button>
      <button onClick={() => onDelete(menu.sheet)} role="menuitem" type="button">
        <span className="menu-icon"><Trash2 aria-hidden="true" /></span>
        <span>删除工作表</span>
        <span />
      </button>
    </div>
  );
}

function ColorPalette({ title, recentColors, allowNone, onSelect }: { title: string; recentColors: readonly string[]; allowNone?: boolean; onSelect: (color: string) => void }) {
  return (
    <div className="format-popover color-palette" role="dialog" aria-label={title} onMouseDown={(event) => event.stopPropagation()}>
      <strong>{title}</strong>
      <span className="palette-section-label">主题颜色</span>
      <div className="palette-grid palette-theme">
        {PALETTE_THEME_ROWS.flat().map((color) => <button className="palette-swatch" key={`${title}-${color}`} onClick={() => onSelect(color)} style={{ backgroundColor: color }} title={color} type="button" />)}
      </div>
      <span className="palette-section-label">标准色</span>
      <div className="palette-grid palette-standard">
        {["#c00000", "#ff0000", "#ffc000", "#ffff00", "#92d050", "#00b050", "#00b0f0", "#0070c0", "#7030a0", "#000000"].map((color) => <button className="palette-swatch" key={`${title}-standard-${color}`} onClick={() => onSelect(color)} style={{ backgroundColor: color }} title={color} type="button" />)}
      </div>
      <span className="palette-section-label">最近使用的颜色</span>
      <div className="palette-grid palette-recent">
        {recentColors.map((color) => <button className="palette-swatch" key={`${title}-recent-${color}`} onClick={() => onSelect(color)} style={{ backgroundColor: color }} title={color} type="button" />)}
      </div>
      {allowNone ? <button className="palette-none-action" onClick={() => onSelect("none")} type="button"><span className="palette-none" />无填充</button> : null}
    </div>
  );
}

const BORDER_ACTIONS: Array<{ action: BorderAction; label: string; shortcut: string; icon: ReactNode }> = [
  { action: "none", label: "无框线", shortcut: "N", icon: <Square aria-hidden="true" /> },
  { action: "all", label: "所有框线", shortcut: "A", icon: <Grid3X3 aria-hidden="true" /> },
  { action: "outside", label: "外侧框线", shortcut: "S", icon: <Square aria-hidden="true" /> },
  { action: "bottom", label: "下框线", shortcut: "O", icon: <AlignEndVertical aria-hidden="true" /> },
  { action: "top", label: "上框线", shortcut: "P", icon: <AlignStartVertical aria-hidden="true" /> },
  { action: "left", label: "左框线", shortcut: "L", icon: <AlignLeft aria-hidden="true" /> },
  { action: "right", label: "右框线", shortcut: "R", icon: <AlignRight aria-hidden="true" /> },
];

function BorderMenu({ onSelect }: { onSelect: (action: BorderAction) => void }) {
  return (
    <div className="border-menu" role="menu" aria-label="边框" onMouseDown={(event) => event.stopPropagation()}>
      <strong>边框</strong>
      {BORDER_ACTIONS.map((item) => (
        <button key={item.action} onClick={() => onSelect(item.action)} role="menuitem" type="button">
          <span className={`border-menu-icon border-menu-icon-${item.action}`}>{item.icon}</span>
          <span>{item.label}</span>
          <kbd>({item.shortcut})</kbd>
        </button>
      ))}
    </div>
  );
}

function cellBorderStyles(
  format: CellFormat,
  above: CellFormat | undefined,
  previous: CellFormat | undefined,
  below: CellFormat | undefined,
  next: CellFormat | undefined,
): CSSProperties {
  const current = format.borders ?? {};
  const top = current.top ?? above?.borders?.bottom;
  const left = current.left ?? previous?.borders?.right;
  const bottom = current.bottom && !below?.borders?.top ? current.bottom : undefined;
  const right = current.right && !next?.borders?.left ? current.right : undefined;
  const result: CSSProperties = {};
  if (top) result.borderTop = `1px solid ${top.color}`;
  if (left) result.borderLeft = `1px solid ${left.color}`;
  if (bottom) result.borderBottom = `1px solid ${bottom.color}`;
  if (right) result.borderRight = `1px solid ${right.color}`;
  return result;
}

function ToolbarSelect<T extends string | number>({
  value,
  options,
  onChange,
  ariaLabel,
  className = "",
}: {
  value: T;
  options: readonly { value: T; label: string }[];
  onChange: (value: T) => void;
  ariaLabel: string;
  className?: string;
}) {
  const [open, setOpen] = useState(false);
  const rootRef = useRef<HTMLDivElement>(null);
  const selected = options.find((option) => option.value === value) ?? options[0];

  useEffect(() => {
    if (!open) return;
    const handleOutside = (event: MouseEvent) => {
      const target = event.target as Node | null;
      if (target && rootRef.current?.contains(target)) return;
      setOpen(false);
    };
    document.addEventListener("mousedown", handleOutside);
    return () => document.removeEventListener("mousedown", handleOutside);
  }, [open]);

  return (
    <div className={`toolbar-select ${open ? "is-open" : ""} ${className}`.trim()} ref={rootRef}>
      <button
        aria-expanded={open}
        aria-haspopup="listbox"
        aria-label={ariaLabel}
        className="toolbar-select-trigger"
        onClick={() => setOpen((current) => !current)}
        type="button"
      >
        <span>{selected?.label ?? ""}</span>
        <ChevronDown aria-hidden="true" />
      </button>
      {open ? (
        <div aria-label={ariaLabel} className="toolbar-select-menu" role="listbox">
          {options.map((option) => (
            <button
              aria-selected={option.value === value}
              className="toolbar-select-option"
              key={String(option.value)}
              onClick={() => { onChange(option.value); setOpen(false); }}
              role="option"
              type="button"
            >
              {option.label}
            </button>
          ))}
        </div>
      ) : null}
    </div>
  );
}

function ActivityFeed({ items }: { items: ActivityItem[] }) {
  return (
    <aside className="activity-panel">
      <div className="activity-heading">
        <div>
          <p className="section-kicker">LIVE ACTIVITY</p>
          <h2>协作动态</h2>
        </div>
        <span className="live-badge"><i />实时</span>
      </div>
      <div className="activity-list">
        {items.length ? (
          items.slice(0, 14).map((item) => (
            <article className="activity-item" key={item.id}>
              <Avatar clientId={item.client_id} name={item.name} size="small" />
              <div>
                <p>
                  <strong>{item.name}</strong>
                  {item.kind === "row" ? " 新增了一行" : " 更新了 "}
                  {item.kind === "cell" ? <span>{item.account_name || `第 ${item.row_id} 行`}</span> : null}
                </p>
                <p className="activity-detail">
                  {item.kind === "row" ? "新增行" : item.field_label}
                  {item.kind === "cell" ? ` · ${formatValue(item.value, typeof item.value === "number" ? "number" : "text")}` : ""}
                </p>
              </div>
              <time dateTime={item.updated_at}>{timeLabel(item.updated_at)}</time>
            </article>
          ))
        ) : (
          <div className="activity-empty">编辑任意单元格后，动态会出现在这里。</div>
        )}
      </div>
      <div className="activity-footnote">
        <span className="shield-icon"><Icon name="check" /></span>
        <p><strong>自动保存已开启</strong><span>每次编辑都会记录贡献人和时间</span></p>
      </div>
    </aside>
  );
}

interface VirtualGridProps {
  fields: FieldDefinition[];
  rows: GridBookRow[];
  state: GridBookState;
  columnWidths: number[];
  selection: CellSelection;
  selections: CellSelection[];
  selectedBounds: ReturnType<typeof selectionBounds>;
  identity: CollaborationIdentity;
  flashes: Set<string>;
  editRequest: { row: number; col: number; id: number; seed: string | null } | null;
  mergeGeometry: { hidden: Set<string>; anchors: Map<string, { rowSpan: number; colSpan: number; id: string }> };
  onScroll: (position: { top: number; left: number; scrollHeight: number; scrollWidth: number; clientHeight: number; clientWidth: number }) => void;
  onCommit: (rowId: number, field: EditableField, value: string | number | null, version: number) => Promise<void>;
  onFocusCell: (cell: { row_id: number; field: EditableField } | null) => void;
  onMove: (rowDelta: number, colDelta: number) => void;
  onSelect: (row: number, col: number, extend: boolean, additive: boolean) => void;
  onSelectRange: (row: number, col: number, rowSpan: number, colSpan: number, extend: boolean, additive: boolean) => void;
  onDrag: (row: number, col: number) => void;
  onFill?: (source: CellSelection, target: CellPosition, copyMode?: boolean, doubleClick?: boolean) => void;
  onFormulaReferencePointerDown?: (row: number, col: number) => boolean;
  onFormulaReferencePointerMove?: (row: number, col: number) => boolean;
  onFormulaEditorChange?: (handle: FormulaEditorHandle | null) => void;
  onFormulaDraftChange?: (value: string) => void;
  formulaReferences?: FormulaReferenceToken[];
  onContext: (event: ReactMouseEvent, row: number, col: number) => void;
  onRowSelect: (row: number, extend: boolean, additive: boolean) => void;
  onColumnSelect: (col: number, extend: boolean, additive: boolean) => void;
  onAll: () => void;
  onColumnResize: (column: number, event: ReactMouseEvent) => void;
  onRowResize: (row: GridBookRow, event: ReactMouseEvent) => void;
  onRowDrag: (row: number) => void;
  onColumnDrag: (col: number) => void;
  cellSelected: (row: number, col: number) => boolean;
  freeze: { row: boolean; column: boolean };
  searchMatches: Set<string>;
  zoom: number;
  formulaContext: { rows: GridBookRow[]; fields: FieldDefinition[] };
}

function VirtualGrid({
  fields, rows, state, columnWidths, selection, selections, selectedBounds, identity, flashes,
  editRequest, mergeGeometry, onScroll, onCommit, onFocusCell, onMove, onSelect,
  onDrag, onFill = () => undefined, onFormulaReferencePointerDown = () => false, onFormulaReferencePointerMove = () => false, onFormulaEditorChange = () => undefined, onFormulaDraftChange = () => undefined, formulaReferences = [], onContext, onRowSelect, onColumnSelect, onAll, onColumnResize, onRowResize, onSelectRange,
  onRowDrag, onColumnDrag, cellSelected, freeze, searchMatches, zoom, formulaContext,
}: VirtualGridProps) {
  const shellRef = useRef<HTMLDivElement>(null);
  const scrollRef = useRef<HTMLDivElement>(null);
  const headTrackRef = useRef<HTMLDivElement>(null);
  const rowTrackRef = useRef<HTMLDivElement>(null);
  const handleScrollRef = useRef<() => void>(() => undefined);
  const [scroll, setScroll] = useState({ top: 0, left: 0 });
  const [viewport, setViewport] = useState({ width: 1180, height: 520 });
  const scrollFrame = useRef<number | null>(null);
  const pendingScroll = useRef(scroll);
  // Keep frozen panes on the compositor path. Rails move as two tracks rather
  // than transforming every label after React has rendered a new tree.
  const syncScrollVisuals = useCallback((target: HTMLDivElement) => {
    const shell = shellRef.current;
    if (!shell) return;
    const scrollX = target.scrollLeft;
    const scrollY = target.scrollTop;
    if (headTrackRef.current) headTrackRef.current.style.transform = `translate3d(${-scrollX}px, 0, 0)`;
    if (rowTrackRef.current) rowTrackRef.current.style.transform = `translate3d(0, ${-scrollY}px, 0)`;
    shell.style.setProperty("--sheet-scroll-x", `${scrollX}px`);
    shell.style.setProperty("--sheet-scroll-y", `${scrollY}px`);
  }, []);
  // Zoom belongs to the sheet viewport, not the typography toolbar. Keep all
  // measured grid geometry in one coordinate system so scrolling, selection,
  // frozen panes and resize handles stay aligned at every zoom level.
  const zoomScale = Math.max(0.5, Math.min(2, zoom / 100));
  const rowHeight = DEFAULT_WORKSHEET_ROW_HEIGHT * zoomScale;
  const headerHeight = 32 * zoomScale;
  const rowNumberWidth = 46 * zoomScale;
  // Keep a generous buffer around the viewport.  The scroll rail moves on
  // the compositor immediately, while React refreshes the virtual window on
  // the next animation frame.  A larger buffer prevents a fast drag from
  // briefly showing recycled row/cell coordinates from the previous window.
  const rowOverscan = Math.max(32, Math.ceil(viewport.height / Math.max(rowHeight, 1)) * 2);
  const columnOverscan = Math.max(32, Math.ceil(viewport.width / 120) * 2);
  // Keep the sheet's row scaffold visible even while the first state request
  // is in flight (and for a genuinely blank workbook). Without this, the
  // virtual grid receives an empty `rows` array and the row-number rail
  // disappears completely until data arrives.
  const renderRowCount = Math.max(rows.length, GRID_ROW_TARGET);
  const rowHeights = useMemo(() => Array.from({ length: renderRowCount }, (_, index) => {
    const row = rows[index];
    const saved = row ? state.layout.row_heights[String(row.id)] : undefined;
    // 55px is the legacy primary-sheet default. Treat it as unset so old
    // persisted layouts do not make the virtual grid inconsistent with the
    // current universal 29px worksheet row height.
    return saved && saved !== LEGACY_DEFAULT_ROW_HEIGHT ? saved * zoomScale : rowHeight;
  }), [renderRowCount, rows, rowHeight, state.layout.row_heights, zoomScale]);
  const rowOffsets = useMemo(() => {
    const offsets = [0];
    for (const height of rowHeights) offsets.push(offsets.at(-1)! + height);
    return offsets;
  }, [rowHeights]);
  const locateOffset = (offsets: number[], value: number) => {
    let low = 0;
    let high = offsets.length - 1;
    while (low < high) {
      const middle = Math.ceil((low + high) / 2);
      if (offsets[middle] <= value) low = middle;
      else high = middle - 1;
    }
    return Math.min(Math.max(low, 0), offsets.length - 2);
  };
  // The horizontal scroll track represents only the columns that currently
  // exist on the sheet. It expands in blocks as the user reaches the edge.
  const visibleColumnWidths = useMemo(
    () => columnWidths.slice(0, fields.length).map((width) => width * zoomScale),
    [columnWidths, fields.length, zoomScale],
  );
  // The scroll viewport starts after the fixed row rail/header.  Its canvas
  // therefore contains only cell geometry; the reserved rails are outside
  // the native scrollbar track.
  const totalWidth = visibleColumnWidths.reduce((sum, width) => sum + width, 0);
  const totalHeight = rowOffsets.at(-1)!;
  const columnOffsets = useMemo(() => {
    const values: number[] = [];
    let offset = 0;
    for (const width of visibleColumnWidths) { values.push(offset); offset += width; }
    return values;
  }, [visibleColumnWidths]);

  // The column-letter strip and row-number rail are the sheet headers. They
  // stay pinned independently from the data canvas. Do not treat rows[0] as a
  // header: it is real user data, so pinning it would duplicate the first data
  // row and make the headers drift out of alignment while scrolling.
  const frozenRowHeight = 0;
  const bodyRowStart = 0;
  const bodyColumnStart = 0;
  const bodyRowBase = 0;
  // A column is never frozen implicitly.  Column freezing will only be
  // introduced together with an explicit user-facing freeze-columns action;
  // keeping the body and header on the same origin prevents A from becoming a
  // surprise frozen pane and keeps horizontal coordinates aligned.
  const bodyColumnBase = 0;
  const bodyHeaderHeight = headerHeight + frozenRowHeight;
  const bodyTotalHeight = Math.max(0, totalHeight - bodyRowBase);
  const bodyTotalWidth = Math.max(0, totalWidth - bodyColumnBase);
  const firstRow = Math.max(bodyRowStart, locateOffset(rowOffsets, bodyRowBase + Math.max(0, scroll.top)) - rowOverscan);
  const lastRow = Math.min(renderRowCount, locateOffset(rowOffsets, bodyRowBase + scroll.top + viewport.height) + rowOverscan + 1);

  // A selection must always contain a complete merged cell.  The old
  // coordinate-only check marked the merge anchor as selected while its
  // rendered rowSpan/colSpan extended outside the blue selection rectangle.
  // That produced the detached light-blue blocks visible below/right of the
  // frame.  Expand intersecting merges until the range is stable, matching
  // the behavior of Excel/WPS/Tencent Docs.
  const normalizeSelectionBounds = useCallback((item: CellSelection) => {
    let bounds = selectionBounds(item);
    let expanded = true;
    while (expanded) {
      expanded = false;
      mergeGeometry.anchors.forEach((merge, key) => {
        const separator = key.indexOf(":");
        if (separator < 0) return;
        const top = Number(key.slice(0, separator));
        const left = Number(key.slice(separator + 1));
        if (!Number.isFinite(top) || !Number.isFinite(left)) return;
        const bottom = top + merge.rowSpan - 1;
        const right = left + merge.colSpan - 1;
        const intersects = !(bounds.bottom < top || bounds.top > bottom || bounds.right < left || bounds.left > right);
        if (!intersects) return;
        const next = {
          top: Math.min(bounds.top, top),
          bottom: Math.max(bounds.bottom, bottom),
          left: Math.min(bounds.left, left),
          right: Math.max(bounds.right, right),
        };
        if (next.top !== bounds.top || next.bottom !== bounds.bottom || next.left !== bounds.left || next.right !== bounds.right) {
          bounds = next;
          expanded = true;
        }
      });
    }
    return bounds;
  }, [mergeGeometry]);
  const normalizedSelections = useMemo(
    () => selections.map((item) => normalizeSelectionBounds(item)),
    [normalizeSelectionBounds, selections],
  );

  const firstColumn = Math.max(bodyColumnStart, locateOffset([...columnOffsets, totalWidth], bodyColumnBase + Math.max(0, scroll.left)) - columnOverscan);
  const lastColumn = Math.min(fields.length, locateOffset([...columnOffsets, totalWidth], bodyColumnBase + scroll.left + viewport.width) + columnOverscan + 1);
  const singleCellSelection = normalizedSelections.length === 1 && normalizedSelections[0].top === normalizedSelections[0].bottom && normalizedSelections[0].left === normalizedSelections[0].right;
  const visibleColumns = useMemo(() => {
    const indices = Array.from({ length: Math.max(0, lastColumn - firstColumn) }, (_, index) => firstColumn + index);
    return indices.filter((columnIndex) => columnIndex >= bodyColumnStart);
  }, [bodyColumnStart, firstColumn, lastColumn]);
  // Column labels are a very small layer compared with the cell body. Keep
  // every label mounted so a large horizontal jump never exposes stale or
  // blank header buttons while the body window is being recycled.
  const visibleHeaderColumns = useMemo(
    () => Array.from({ length: fields.length }, (_, index) => index),
    [fields.length],
  );
  const visibleRows = useMemo(() => {
    const indices = Array.from({ length: Math.max(0, lastRow - firstRow) }, (_, index) => firstRow + index);
    return indices.filter((rowIndex) => rowIndex >= bodyRowStart);
  }, [bodyRowStart, firstRow, lastRow]);

  // The fill handle is intentionally implemented outside individual cells so
  // it remains stable while the virtualized row window is recycled.
  const [fillTarget, setFillTarget] = useState<CellPosition | null>(null);
  const fillDragRef = useRef<{ source: CellSelection; copyMode: boolean } | null>(null);
  const fillTargetRef = useRef<CellPosition | null>(null);
  const fillPointerRef = useRef<{ x: number; y: number } | null>(null);
  const fillAutoScrollFrame = useRef<number | null>(null);
  const [fillMode, setFillMode] = useState<FillMode>("auto");
  const pointToCell = useCallback((clientX: number, clientY: number): CellPosition | null => {
    const element = scrollRef.current;
    if (!element || !fields.length || !rows.length) return null;
    const rect = element.getBoundingClientRect();
    const x = clientX - rect.left + element.scrollLeft + bodyColumnBase;
    const y = clientY - rect.top + element.scrollTop + bodyRowBase;
    if (y < 0) return null;
    const column = Math.max(0, Math.min(fields.length - 1, locateOffset([...columnOffsets, totalWidth], x)));
    const row = Math.max(0, Math.min(rows.length - 1, locateOffset(rowOffsets, y)));
    return { row, col: column };
  }, [bodyColumnBase, bodyRowBase, columnOffsets, fields.length, rowOffsets, rows.length, totalWidth]);

  const updateFillTarget = useCallback((clientX: number, clientY: number) => {
    const element = scrollRef.current;
    if (!element || !fillDragRef.current) return;
    const rect = element.getBoundingClientRect();
    const edge = 42;
    const speed = 26;
    const scrollX = clientX < rect.left + edge ? -speed : clientX > rect.right - edge ? speed : 0;
    const scrollY = clientY < rect.top + edge ? -speed : clientY > rect.bottom - edge ? speed : 0;
    if (scrollX || scrollY) {
      element.scrollLeft = Math.max(0, element.scrollLeft + scrollX);
      element.scrollTop = Math.max(0, element.scrollTop + scrollY);
    }
    const next = pointToCell(clientX, clientY);
    if (!next) return;
    fillTargetRef.current = next;
    setFillTarget(next);
  }, [pointToCell]);

  useEffect(() => {
    const tick = () => {
      fillAutoScrollFrame.current = null;
      const pointer = fillPointerRef.current;
      if (!fillDragRef.current || !pointer) return;
      updateFillTarget(pointer.x, pointer.y);
      fillAutoScrollFrame.current = requestAnimationFrame(tick);
    };
    const onMove = (event: PointerEvent) => {
      if (!fillDragRef.current) return;
      fillPointerRef.current = { x: event.clientX, y: event.clientY };
      updateFillTarget(event.clientX, event.clientY);
      if (fillAutoScrollFrame.current === null) fillAutoScrollFrame.current = requestAnimationFrame(tick);
    };
    const onUp = () => {
      const drag = fillDragRef.current;
      const target = fillTargetRef.current;
      if (drag && target) onFill(drag.source, target, drag.copyMode);
      fillDragRef.current = null;
      fillTargetRef.current = null;
      fillPointerRef.current = null;
      setFillTarget(null);
      setFillMode("auto");
      if (fillAutoScrollFrame.current !== null) {
        window.cancelAnimationFrame(fillAutoScrollFrame.current);
        fillAutoScrollFrame.current = null;
      }
    };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
    window.addEventListener("pointercancel", onUp);
    return () => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      window.removeEventListener("pointercancel", onUp);
      if (fillAutoScrollFrame.current !== null) window.cancelAnimationFrame(fillAutoScrollFrame.current);
    };
  }, [onFill, updateFillTarget]);

  const sourceBounds = selectionBounds(selection);
  const fillBounds = fillTarget ? {
    top: Math.min(sourceBounds.top, fillTarget.row),
    bottom: Math.max(sourceBounds.bottom, fillTarget.row),
    left: Math.min(sourceBounds.left, fillTarget.col),
    right: Math.max(sourceBounds.right, fillTarget.col),
  } : null;
  const fillPreview = useMemo(() => {
    if (!fillBounds || !fillTarget) return new Map<string, ReturnType<typeof fillCellAt>>();
    const grid = { rows, fields, formats: state.layout.formats };
    return buildFillPreview(grid, sourceBounds, fillBounds, fillMode);
  }, [fields, fillBounds, fillMode, fillTarget, rows, sourceBounds, state.layout.formats]);
  const startColumn = Math.max(0, Math.min(fields.length - 1, fillBounds?.left ?? sourceBounds.left));
  const endColumn = Math.max(startColumn, Math.min(fields.length - 1, fillBounds?.right ?? sourceBounds.right));
  const startRow = Math.max(0, Math.min(rows.length - 1, fillBounds?.top ?? sourceBounds.top));
  const endRow = Math.max(startRow, Math.min(rows.length - 1, fillBounds?.bottom ?? sourceBounds.bottom));
  const fillLeft = columnOffsets[startColumn] ?? 0;
  const fillWidth = visibleColumnWidths.slice(startColumn, endColumn + 1).reduce((sum, value) => sum + value, 0);
  const fillTop = rowOffsets[startRow] ?? 0;
  const fillHeight = (rowOffsets[endRow + 1] ?? fillTop) - fillTop;
  const handleLeft = (columnOffsets[sourceBounds.right + 1] ?? fillLeft + fillWidth) - 3;
  const handleTop = (rowOffsets[sourceBounds.bottom + 1] ?? fillTop + fillHeight) - 3;

  useEffect(() => {
    const element = scrollRef.current;
    if (!element) return undefined;
    const update = () => setViewport({ width: element.clientWidth, height: element.clientHeight });
    update();
    const observer = new ResizeObserver(update);
    observer.observe(element);
    return () => observer.disconnect();
  }, []);

  useEffect(() => {
    const element = scrollRef.current;
    if (element) syncScrollVisuals(element);
  // The virtualized tracks can be replaced when the visible window changes.
  // Re-apply the compositor offsets after that replacement so frozen headers,
  // row numbers and the frozen first column never fall back to stale origins.
  }, [bodyColumnBase, bodyHeaderHeight, freeze.row, scroll.left, scroll.top, syncScrollVisuals]);

  useEffect(() => () => {
    if (scrollFrame.current !== null) window.cancelAnimationFrame(scrollFrame.current);
  }, []);

  const handleScroll = () => {
    const target = scrollRef.current;
    if (!target) return;
    const nextScroll = { top: target.scrollTop, left: target.scrollLeft };
    pendingScroll.current = nextScroll;
    // The passive native listener moves frozen rails immediately. React only
    // recalculates the virtualized window when the visible range changes.
    if (scrollFrame.current === null) {
      scrollFrame.current = requestAnimationFrame(() => {
        scrollFrame.current = null;
        const next = pendingScroll.current;
        const nextFirstRow = Math.max(bodyRowStart, locateOffset(rowOffsets, bodyRowBase + Math.max(0, next.top)) - rowOverscan);
        const nextLastRow = Math.min(renderRowCount, locateOffset(rowOffsets, bodyRowBase + next.top + viewport.height) + rowOverscan + 1);
        const nextFirstColumn = Math.max(bodyColumnStart, locateOffset([...columnOffsets, totalWidth], bodyColumnBase + Math.max(0, next.left)) - columnOverscan);
        const nextLastColumn = Math.min(fields.length, locateOffset([...columnOffsets, totalWidth], bodyColumnBase + next.left + viewport.width) + columnOverscan + 1);
        if (nextFirstRow !== firstRow || nextLastRow !== lastRow || nextFirstColumn !== firstColumn || nextLastColumn !== lastColumn) setScroll(next);
        const current = scrollRef.current;
        if (current) {
          // Expanding the sparse sheet is the only reason the parent needs a
          // scroll notification. Do not re-render the whole workbook for
          // every animation-frame while the user is dragging the scrollbar.
          // The browser and compositor tracks already handle the visual move.
          const nearBottom = current.scrollHeight - next.top - current.clientHeight < 850;
          const nearRight = current.scrollWidth - next.left - current.clientWidth < Math.max(current.clientWidth, 1200);
          if (nearBottom || nearRight) {
            onScroll({
              top: next.top,
              left: next.left,
              scrollHeight: current.scrollHeight,
              scrollWidth: current.scrollWidth,
              clientHeight: current.clientHeight,
              clientWidth: current.clientWidth,
            });
          }
        }
      });
    }
  };
  handleScrollRef.current = handleScroll;

  // Keep scroll input off React's synthetic event path. The native passive
  // listener lets the browser move the sheet and frozen rails immediately,
  // while React only updates the virtual window once per animation frame.
  useEffect(() => {
    const element = scrollRef.current;
    if (!element) return undefined;
    const update = () => {
      syncScrollVisuals(element);
      handleScrollRef.current();
    };
    element.addEventListener("scroll", update, { passive: true });
    update();
    return () => element.removeEventListener("scroll", update);
  }, [syncScrollVisuals]);

  const renderGridCell = (rowIndex: number, columnIndex: number, left: number, top: number, extraClass = "") => {
    const row = rows[rowIndex];
    const field = fields[columnIndex];
    if (!field) return null;
    const currentRowHeight = rowHeights[rowIndex] ?? rowHeight;
    const coordinateKey = `${rowIndex}:${columnIndex}`;
    if (mergeGeometry.hidden.has(coordinateKey)) return null;
    const mergeInfo = mergeGeometry.anchors.get(coordinateKey);
    const width = visibleColumnWidths.slice(columnIndex, columnIndex + (mergeInfo?.colSpan ?? 1)).reduce((sum, value) => sum + value, 0);
    const mergeHeight = rowOffsets[Math.min(renderRowCount, rowIndex + (mergeInfo?.rowSpan ?? 1))] - rowOffsets[rowIndex];
    const format = row ? state.layout.formats[`${row.id}:${field.key}`] ?? {} : {};
    const isSelected = normalizedSelections.some((bounds) => rowIndex >= bounds.top && rowIndex <= bounds.bottom && columnIndex >= bounds.left && columnIndex <= bounds.right);
    const isActive = singleCellSelection && selection.focus.row === rowIndex && selection.focus.col === columnIndex;
    const cellBottom = rowIndex + (mergeInfo?.rowSpan ?? 1) - 1;
    const cellRight = columnIndex + (mergeInfo?.colSpan ?? 1) - 1;
    const key = row ? `${row.id}:${field.key}` : coordinateKey;
    const remoteEditor = row ? state.presence.find((person) => person.client_id !== identity.clientId && person.cell?.row_id === row.id && person.cell.field === field.key) : undefined;
    const fillPreviewCell = fillPreview.get(coordinateKey);
    const renderedFormat = fillPreviewCell?.format ?? format;
    const formulaReference = formulaReferences.find((reference) => rowIndex >= reference.startCell.row && rowIndex <= reference.endCell.row && columnIndex >= reference.startCell.col && columnIndex <= reference.endCell.col);
    const aboveFormat = rowIndex > 0 ? state.layout.formats[`${rows[rowIndex - 1]?.id}:${field.key}`] : undefined;
    const previousFormat = columnIndex > 0 && row ? state.layout.formats[`${row.id}:${fields[columnIndex - 1]?.key}`] : undefined;
    const belowFormat = rowIndex + 1 < rows.length ? state.layout.formats[`${rows[rowIndex + 1]?.id}:${field.key}`] : undefined;
    const nextFormat = row && columnIndex + 1 < fields.length ? state.layout.formats[`${row.id}:${fields[columnIndex + 1]?.key}`] : undefined;
    const borderStyles = cellBorderStyles(format, aboveFormat, previousFormat, belowFormat, nextFormat);
    const firstRowRemote = rowIndex === 0 && Boolean(remoteEditor);
    const className = ["grid-cell", extraClass, mergeInfo ? "merged-cell" : "", fillPreviewCell ? "fill-preview-cell" : "", firstRowRemote ? "remote-editor-anchor" : "", isSelected ? "cell-selected" : "", isActive ? "cell-active" : "", formulaReference ? "formula-reference-cell" : "", searchMatches.has(coordinateKey) ? "search-match" : "", renderedFormat.bold ? "format-bold" : "", renderedFormat.italic ? "format-italic" : "", renderedFormat.underline ? "format-underline" : "", renderedFormat.wrap ? "format-wrap" : "", renderedFormat.align ? `format-align-${renderedFormat.align}` : "", renderedFormat.vertical_align ? `format-vertical-${renderedFormat.vertical_align}` : "", renderedFormat.border && !renderedFormat.borders ? "format-border" : ""].filter(Boolean).join(" ");
    const style = { left, top, width, height: mergeHeight || currentRowHeight, fontSize: renderedFormat.font_size ? `calc(${renderedFormat.font_size}px * var(--sheet-zoom, 1))` : undefined, color: renderedFormat.font_color && renderedFormat.font_color !== "default" ? renderedFormat.font_color : undefined, backgroundColor: renderedFormat.fill && renderedFormat.fill !== "none" ? renderedFormat.fill : undefined, ...borderStyles, ...(firstRowRemote && remoteEditor ? avatarStyle(remoteEditor.client_id) : {}), ...(formulaReference ? { "--formula-reference-color": formulaReference.color, "--formula-reference-fill": formulaReference.fill } : {}) } as CSSProperties;
    if (!row) return <div className={`${className} grid-cell-placeholder`} key={`placeholder-cell-${rowIndex}-${columnIndex}`} style={style} />;
    return <div className={className} key={field.key} style={style}>
      <EditableCell displayValue={fillPreviewCell ? fillPreviewCell.value : displayFormulaValue(row[field.key], formulaContext)} field={field} flash={flashes.has(key)} format={renderedFormat} identity={identity} meta={state.cells[key]} active={false} editRequest={editRequest?.row === rowIndex && editRequest.col === columnIndex ? { id: editRequest.id, seed: editRequest.seed } : undefined} onCommit={(value, version) => onCommit(row.id, field.key, value, version)} onFocusCell={onFocusCell} onFormulaDraftChange={onFormulaDraftChange} onFormulaEditorChange={onFormulaEditorChange} onFormulaReferencePointerDown={() => onFormulaReferencePointerDown(rowIndex, columnIndex)} onFormulaReferencePointerMove={() => onFormulaReferencePointerMove(rowIndex, columnIndex)} onMoveAfterCommit={onMove} onOpenContextMenu={(event) => onContext(event, rowIndex, columnIndex)} onSelect={(extend, additive) => mergeInfo ? onSelectRange(rowIndex, columnIndex, mergeInfo.rowSpan, mergeInfo.colSpan, extend, additive) : onSelect(rowIndex, columnIndex, extend, additive)} onSelectionDrag={() => onDrag(cellBottom, cellRight)} remoteEditor={remoteEditor} remoteEditorInHeader={rowIndex === 0} row={row} selected={isSelected} selectionEdges={{ top: false, right: false, bottom: false, left: false }} />
    </div>;
  };

  return (
    <div ref={shellRef} className="virtual-grid-shell" style={{ "--sheet-zoom": String(zoomScale) } as CSSProperties}>
      <div className="virtual-grid-fixed-head" style={{ height: bodyHeaderHeight }}>
        <button className="grid-corner" aria-label="全选表格" title="全选表格" style={{ left: 0, width: rowNumberWidth }} onMouseDown={(event) => { event.preventDefault(); onAll(); }} type="button" />
        <div className="virtual-grid-fixed-head-viewport" style={{ left: rowNumberWidth }}>
        <div ref={headTrackRef} className="virtual-grid-fixed-head-track" style={{ width: totalWidth, height: headerHeight }}>
        {visibleHeaderColumns.map((columnIndex) => {
          const field = fields[columnIndex];
          const left = columnOffsets[columnIndex] ?? 0;
          const firstRowRemoteEditor = rows[0]
            ? state.presence.find(
              (person) =>
                person.client_id !== identity.clientId &&
                person.cell?.row_id === rows[0].id &&
                person.cell.field === field.key,
            )
            : undefined;
          return <Fragment key={`fixed-head-${field.key}`}>
            <button className={normalizedSelections.some((bounds) => columnIndex >= bounds.left && columnIndex <= bounds.right) ? "grid-col-head selected" : "grid-col-head"} style={{ left, width: visibleColumnWidths[columnIndex] }} onMouseDown={(event) => { event.preventDefault(); onColumnSelect(columnIndex, event.shiftKey, event.ctrlKey || event.metaKey); }} onMouseEnter={() => onColumnDrag(columnIndex)} type="button">
              <span>{excelColumnName(columnIndex)}</span><i onMouseDown={(event) => onColumnResize(columnIndex, event)} />
            </button>
            {firstRowRemoteEditor ? (
              <span
                aria-label={`${firstRowRemoteEditor.name} 正在输入`}
                className="remote-editor remote-editor-fixed-head"
                style={{ ...avatarStyle(firstRowRemoteEditor.client_id), left }}
              >
                {firstRowRemoteEditor.name} 正在输入
              </span>
            ) : null}
          </Fragment>;
        })}
        </div>
        </div>
      </div>
      <div className="virtual-grid-fixed-rows" style={{ top: headerHeight, width: rowNumberWidth, bottom: 9 }}>
        <div ref={rowTrackRef} className="virtual-grid-fixed-rows-track" style={{ width: rowNumberWidth, height: bodyTotalHeight, top: frozenRowHeight }}>
        {visibleRows.map((rowIndex) => {
          const row = rows[rowIndex];
          const top = (rowOffsets[rowIndex] ?? 0) - bodyRowBase;
          return <button className={normalizedSelections.some((bounds) => rowIndex >= bounds.top && rowIndex <= bounds.bottom) ? "grid-row-head selected" : "grid-row-head"} key={`fixed-row-${row?.id ?? `placeholder-${rowIndex}`}`} style={{ top, width: rowNumberWidth, height: rowHeights[rowIndex] }} onMouseDown={row ? (event) => { event.preventDefault(); onRowSelect(rowIndex, event.shiftKey, event.ctrlKey || event.metaKey); } : undefined} onMouseEnter={row ? () => onRowDrag(rowIndex) : undefined} type="button">{rowIndex + 1}{row && <i onMouseDown={(event) => onRowResize(row, event)} />}</button>;
        })}
        </div>
      </div>
      <div className="virtual-grid-scroll" ref={scrollRef} style={{ top: bodyHeaderHeight, left: rowNumberWidth }}>
        <div className="virtual-grid-canvas" style={{ width: bodyTotalWidth, height: bodyTotalHeight }}>
          <div className="virtual-grid-body" style={{ width: bodyTotalWidth }}>
            {visibleRows.map((rowIndex) => {
              const row = rows[rowIndex];
              const currentRowHeight = rowHeights[rowIndex];
              if (!row) {
                return <div className="grid-row grid-row-placeholder" key={`placeholder-row-${rowIndex}`} style={{ width: bodyTotalWidth, top: (rowOffsets[rowIndex] ?? 0) - bodyRowBase, height: currentRowHeight }}>
                  {visibleColumns.map((columnIndex) => renderGridCell(rowIndex, columnIndex, (columnOffsets[columnIndex] ?? 0) - bodyColumnBase, 0))}
                </div>;
              }
              return (
                <div className="grid-row" key={row.id} style={{ width: bodyTotalWidth, top: (rowOffsets[rowIndex] ?? 0) - bodyRowBase, height: currentRowHeight }}>
                  {visibleColumns.map((columnIndex) => renderGridCell(rowIndex, columnIndex, (columnOffsets[columnIndex] ?? 0) - bodyColumnBase, 0))}
                </div>
              );
            })}
            <div
              className="grid-selection-layer"
              aria-hidden="true"
              style={{ width: bodyTotalWidth, height: bodyTotalHeight }}
            >
              {normalizedSelections.map((bounds, index) => {
                const startColumn = Math.max(0, Math.min(fields.length - 1, bounds.left));
                const endColumn = Math.max(startColumn, Math.min(fields.length - 1, bounds.right));
                const startRow = Math.max(0, Math.min(rows.length - 1, bounds.top));
                const endRow = Math.max(startRow, Math.min(rows.length - 1, bounds.bottom));
                const left = Math.max(0, (columnOffsets[startColumn] ?? 0) - bodyColumnBase);
                const width = visibleColumnWidths.slice(startColumn, endColumn + 1).reduce((sum, widthValue) => sum + widthValue, 0);
                const top = Math.max(0, (rowOffsets[startRow] ?? 0) - bodyRowBase);
                const height = (rowOffsets[endRow + 1] ?? top) - top;
                const single = bounds.top === bounds.bottom && bounds.left === bounds.right;
                return <div className={`selection-range-overlay${single ? " selection-range-overlay-single" : ""}`} key={`selection-overlay-${index}`} style={{ left, top, width, height }} />;
              })}
            </div>
            {formulaReferences.map((reference, index) => {
              const startColumn = Math.max(0, Math.min(fields.length - 1, reference.startCell.col));
              const endColumn = Math.max(startColumn, Math.min(fields.length - 1, reference.endCell.col));
              const startRow = Math.max(0, Math.min(rows.length - 1, reference.startCell.row));
              const endRow = Math.max(startRow, Math.min(rows.length - 1, reference.endCell.row));
              const left = Math.max(0, (columnOffsets[startColumn] ?? 0) - bodyColumnBase);
              const top = Math.max(0, (rowOffsets[startRow] ?? 0) - bodyRowBase);
              const width = visibleColumnWidths.slice(startColumn, endColumn + 1).reduce((sum, widthValue) => sum + widthValue, 0);
              const height = (rowOffsets[endRow + 1] ?? top) - (rowOffsets[startRow] ?? top);
              return <div className="formula-reference-overlay" key={`formula-reference-overlay-${index}`} style={{ left, top, width, height, borderColor: reference.color, background: reference.fill }} />;
            })}
            {fillTarget && fillBounds && (
              <div className="fill-preview-overlay" style={{ left: fillLeft, top: fillTop, width: fillWidth, height: fillHeight }} />
            )}
            {rows.length > 0 && selections.length === 1 && (
              <button
                className="fill-handle"
                aria-label="填充柄"
                type="button"
                style={{ left: handleLeft, top: handleTop }}
                onPointerDown={(event) => {
                  event.preventDefault();
                  event.stopPropagation();
                  event.currentTarget.setPointerCapture(event.pointerId);
                  const source = { anchor: { ...selection.anchor }, focus: { ...selection.focus } };
                  const target = { row: sourceBounds.bottom, col: sourceBounds.right };
                  const copyMode = event.ctrlKey || event.metaKey;
                  fillDragRef.current = { source, copyMode };
                  fillTargetRef.current = target;
                  setFillMode(copyMode ? "copy" : "auto");
                  setFillTarget(target);
                }}
                onDoubleClick={(event) => {
                  event.preventDefault();
                  event.stopPropagation();
                  const source = { anchor: { ...selection.anchor }, focus: { ...selection.focus } };
                  onFill(source, { row: sourceBounds.bottom, col: sourceBounds.right }, false, true);
                }}
              />
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

function LegacyApp() {
  const [identity, setIdentity] = useState<CollaborationIdentity>(() => ({
    clientId: createClientId(),
    name: localStorage.getItem(NAME_KEY) ?? "用户",
  }));
  const [identityOpen, setIdentityOpen] = useState(false);
  const [state, setState] = useState<GridBookState>(EMPTY_STATE);
  const [displayedRowCount, setDisplayedRowCount] = useState(GRID_ROW_TARGET);
  const [displayedColumnCount, setDisplayedColumnCount] = useState(30);
  const [loading, setLoading] = useState(true);
  const loadedRowLimitRef = useRef(INITIAL_STATE_LIMIT);
  const gridEnsureInFlightRef = useRef(0);
  // A refresh can overlap an optimistic layout mutation. Keep a request
  // sequence and pending format operations so an older snapshot cannot erase
  // a local edit while its server acknowledgement is still in flight.
  const refreshSequenceRef = useRef(0);
  const pendingFormatOperationsRef = useRef<Record<number, Record<string, CellFormat>>>({});
  const formatOperationIdRef = useRef(0);
  const [connection, setConnection] = useState<"connecting" | "live" | "reconnecting">("connecting");
  const [focusedCell, setFocusedCell] = useState<{ row_id: number; field: EditableField } | null>(null);
  const [flashes, setFlashes] = useState<Set<string>>(new Set());
  const [toast, setToast] = useState<string | null>(null);
  const [addingRow, setAddingRow] = useState(false);
  const [pasting, setPasting] = useState(false);
  const [selection, setSelection] = useState<CellSelection>({
    anchor: { row: 0, col: 0 },
    focus: { row: 0, col: 0 },
  });
  const [extraSelections, setExtraSelections] = useState<CellSelection[]>([]);
  // A fresh workbook must start from the same neutral spreadsheet geometry as
  // every other sheet.  Explicitly saved column sizes still win below.
  const [columnWidths, setColumnWidths] = useState(() => FIELDS.map(() => DEFAULT_WORKSHEET_COLUMN_WIDTH));
  const [editRequest, setEditRequest] = useState<{
    row: number;
    col: number;
    id: number;
    seed: string | null;
  } | null>(null);
  const [contextMenu, setContextMenu] = useState<ContextMenuState | null>(null);
  const [sheetContextMenu, setSheetContextMenu] = useState<SheetContextMenuState | null>(null);
  const [formatBrush, setFormatBrush] = useState<CellFormat | null>(null);
  const [formatBrushContinuous, setFormatBrushContinuous] = useState(false);
  const [borderMenuOpen, setBorderMenuOpen] = useState(false);
  const [formatPalette, setFormatPalette] = useState<"fill" | "font" | null>(null);
  const [formatPanelOpen, setFormatPanelOpen] = useState(false);
  const [formatPanelPosition, setFormatPanelPosition] = useState({ top: 0, left: 0 });
  const [findPanelOpen, setFindPanelOpen] = useState(false);
  const [findPanelPosition, setFindPanelPosition] = useState({ top: 0, left: 0 });
  const [findQuery, setFindQuery] = useState("");
  const [replaceText, setReplaceText] = useState("");
  const [findCursor, setFindCursor] = useState(0);
  const [freeze, setFreeze] = useState({ row: false, column: false });
  const [sheetNames, setSheetNames] = useState<string[]>(() => {
    try {
      const saved = JSON.parse(localStorage.getItem(SHEET_NAMES_STORAGE_KEY) ?? "null");
      return Array.isArray(saved) && saved.every((name) => typeof name === "string") && saved.length
        ? Array.from(new Set([PRIMARY_SHEET_NAME, ...saved]))
        : [...DEFAULT_SHEET_NAMES];
    } catch {
      return [...DEFAULT_SHEET_NAMES];
    }
  });
  const [activeSheet, setActiveSheet] = useState(PRIMARY_SHEET_NAME);
  const [editingSheetName, setEditingSheetName] = useState<string | null>(null);
  const [sheetNameDraft, setSheetNameDraft] = useState("");
  const [formulaDraft, setFormulaDraft] = useState("");
  const [cellFormulaDraft, setCellFormulaDraft] = useState("");
  const [formulaBarFocused, setFormulaBarFocused] = useState(false);
  const formulaDraftRef = useRef("");
  const [zoom, setZoom] = useState(100);
  const [menuOpen, setMenuOpen] = useState(false);
  const toastTimer = useRef<number | null>(null);
  const activeSheetRef = useRef(activeSheet);
  const sheetStatesRef = useRef<Record<string, GridBookState>>({ [PRIMARY_SHEET_NAME]: EMPTY_STATE });
  activeSheetRef.current = activeSheet;
  const tableCardRef = useRef<HTMLElement>(null);
  const formulaInputRef = useRef<HTMLInputElement>(null);
  const cellFormulaEditorRef = useRef<FormulaEditorHandle | null>(null);
  const formulaBarEditorRef = useRef<FormulaEditorHandle | null>(null);
  const formulaReferenceDragRef = useRef<{ editor: FormulaEditorHandle; start: CellPosition } | null>(null);
  const formulaBarReferenceBaseRef = useRef<{ prefix: string; suffix: string } | null>(null);
  const draggingSelection = useRef(false);
  const selectionDragTarget = useRef<CellPosition | null>(null);
  const selectionDragFrame = useRef<number | null>(null);
  const editRequestId = useRef(0);
  const formatBrushSource = useRef("");
  const historyUndoStack = useRef<HistoryAction[]>([]);
  const historyRedoStack = useRef<HistoryAction[]>([]);
  const formatUndoStack = useRef<FormatHistoryItem[]>([]);
  const formatRedoStack = useRef<FormatHistoryItem[]>([]);
  const sheetContextPointRef = useRef({ x: 0, y: 0 });

  const getFlyoutPosition = useCallback((event: ReactMouseEvent<HTMLButtonElement>, width: number, height: number) => {
    const rect = event.currentTarget.getBoundingClientRect();
    const margin = 10;
    const maxLeft = Math.max(margin, window.innerWidth - width - margin);
    const maxTop = Math.max(margin, window.innerHeight - height - margin);
    return {
      left: Math.min(Math.max(margin, rect.left), maxLeft),
      top: Math.min(Math.max(margin, rect.bottom + 6), maxTop),
    };
  }, []);

  // Tool palettes and the document menu behave like Excel flyouts: clicking
  // anywhere outside the active flyout dismisses it. Find/replace is a
  // persistent panel and is intentionally excluded from this handler.
  useEffect(() => {
    const handleOutsidePointer = (event: MouseEvent) => {
      const target = event.target as HTMLElement | null;
      if (!target) return;

      const inPalette = Boolean(target.closest(".color-palette, .toolbar-menu-anchor"))
        || Boolean(target.closest('button[title="填充"]'))
        || Boolean(target.closest('button[title="字色"]'));
      if (formatPalette && !inPalette) setFormatPalette(null);

      const inFormatPanel = Boolean(target.closest(".format-panel, .format-panel-trigger"))
        || Boolean(target.closest('button[title="格式"]'));
      if (formatPanelOpen && !inFormatPanel) setFormatPanelOpen(false);

      const inBorderMenu = Boolean(target.closest(".border-menu, .border-menu-trigger"));
      if (borderMenuOpen && !inBorderMenu) setBorderMenuOpen(false);

      const inDocumentMenu = Boolean(target.closest(".document-menu"))
        || Boolean(target.closest('button[aria-label="菜单"]'));
      if (menuOpen && !inDocumentMenu) setMenuOpen(false);
    };

    document.addEventListener("mousedown", handleOutsidePointer);
    return () => document.removeEventListener("mousedown", handleOutsidePointer);
  }, [borderMenuOpen, formatPalette, formatPanelOpen, menuOpen]);

  const showToast = useCallback((message: string) => {
    setToast(message);
    if (toastTimer.current) window.clearTimeout(toastTimer.current);
    toastTimer.current = window.setTimeout(() => setToast(null), 3200);
  }, []);

  useEffect(() => {
    const closeSheetMenu = (event: MouseEvent) => {
      // A right-click opens the menu; do not immediately close it from the
      // document-level mousedown listener that handles ordinary clicks.
      if (event.button === 2) return;
      const target = event.target as HTMLElement | null;
      if (target?.closest(".sheet-context-menu")) return;
      setSheetContextMenu(null);
    };
    // Close only after a normal click.  Using mousedown here races the
    // browser's contextmenu event on right-click and can clear the menu
    // immediately after the tab handler opens it.
    document.addEventListener("click", closeSheetMenu);
    return () => {
      document.removeEventListener("click", closeSheetMenu);
    };
  }, []);

  useEffect(() => {
    const captureSheetContextPoint = (event: MouseEvent) => {
      const target = (event.target as HTMLElement | null)?.closest(".sheet-tab");
      if (!target) return;
      sheetContextPointRef.current = { x: event.clientX, y: event.clientY };
    };
    document.addEventListener("contextmenu", captureSheetContextPoint, true);
    return () => document.removeEventListener("contextmenu", captureSheetContextPoint, true);
  }, []);

  useEffect(() => {
    sheetStatesRef.current[activeSheet] = state;
  }, [activeSheet, state]);

  useEffect(() => {
    localStorage.setItem(SHEET_NAMES_STORAGE_KEY, JSON.stringify(sheetNames));
  }, [sheetNames]);

  const contentColumnCount = useMemo(() => {
    let last = INITIAL_DATA_FIELD_COUNT;
    state.rows.forEach((row) => {
      FIELDS.forEach((field, index) => {
        if (row[field.key] !== null && row[field.key] !== undefined && row[field.key] !== "") last = Math.max(last, index + 1);
      });
    });
    return last;
  }, [state.rows]);
  const activeFields = useMemo(() => FIELDS.slice(0, Math.max(displayedColumnCount, contentColumnCount)), [contentColumnCount, displayedColumnCount]);
  const currentCell = useMemo(() => {
    const row = state.rows[selection.focus.row];
    const field = FIELDS[selection.focus.col];
    if (!row || !field) return null;
    return { row, field, value: row[field.key] };
  }, [selection.focus.col, selection.focus.row, state.rows]);
  // Keep the formula bar backed by the raw value of the focused cell.  The
  // grid may render a formatted/calculated display value, but the formula bar
  // must always expose the complete editable content (including long text).
  const currentCellValue = currentCell?.value === null || currentCell?.value === undefined
    ? ""
    : String(currentCell.value);
  const currentCellName = `${excelColumnName(selection.focus.col)}${selection.focus.row + 1}`;
  formulaDraftRef.current = formulaDraft;
  const formulaReferenceText = formulaBarFocused ? formulaDraft : cellFormulaDraft;
  const formulaReferences = useMemo(() => parseFormulaReferences(formulaReferenceText), [formulaReferenceText]);

  useEffect(() => {
    setFormulaDraft(currentCellValue);
  }, [currentCellValue, selection.focus.col, selection.focus.row]);

  const refresh = useCallback(async (limit = loadedRowLimitRef.current) => {
    const requestId = ++refreshSequenceRef.current;
    try {
      const next = await getState(limit);
      if (activeSheetRef.current !== PRIMARY_SHEET_NAME || requestId !== refreshSequenceRef.current) return;
      setState((current) => {
        const shouldReplace = current.rows.length === 0 || next.revision > current.revision;
        const pendingFormats = Object.values(pendingFormatOperationsRef.current)
          .flatMap((operation) => Object.entries(operation));
        if (!shouldReplace && !pendingFormats.length) return current;

        const base = shouldReplace ? next : current;
        if (!pendingFormats.length) {
          sheetStatesRef.current[PRIMARY_SHEET_NAME] = base;
          return base;
        }

        const formats = { ...base.layout.formats };
        for (const [key, format] of pendingFormats) {
          formats[key] = { ...(formats[key] ?? {}), ...format };
        }
        const merged = { ...base, layout: { ...base.layout, formats } };
        sheetStatesRef.current[PRIMARY_SHEET_NAME] = merged;
        return merged;
      });
    } catch (error) {
      showToast(error instanceof Error ? error.message : "暂时无法加载协作表");
    } finally {
      setLoading(false);
    }
  }, [showToast]);

  const switchSheet = useCallback((sheetName: string) => {
    if (sheetName === activeSheet) return;
    sheetStatesRef.current[activeSheet] = state;
    const nextState = sheetStatesRef.current[sheetName] ?? createBlankSheetState();
    sheetStatesRef.current[sheetName] = nextState;
    setActiveSheet(sheetName);
    setState(nextState);
    setSelection({ anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } });
    setExtraSelections([]);
    setEditRequest(null);
    setFocusedCell(null);
    setFindQuery("");
    setFindCursor(0);
    setFormatPalette(null);
    setFormatPanelOpen(false);
    setFindPanelOpen(false);
    setDisplayedRowCount(Math.max(GRID_ROW_TARGET, nextState.rows.length));
    setDisplayedColumnCount(Math.max(30, INITIAL_DATA_FIELD_COUNT));
    setLoading(sheetName === PRIMARY_SHEET_NAME ? loading : false);
    if (sheetName === PRIMARY_SHEET_NAME) {
      loadedRowLimitRef.current = INITIAL_STATE_LIMIT;
      void refresh(INITIAL_STATE_LIMIT);
    }
  }, [activeSheet, loading, refresh, state]);

  const addWorksheet = useCallback(() => {
    let index = sheetNames.length + 1;
    let name = `工作表${index}`;
    while (sheetNames.includes(name)) {
      index += 1;
      name = `工作表${index}`;
    }
    const blankState = createBlankSheetState();
    sheetStatesRef.current[activeSheet] = state;
    sheetStatesRef.current[name] = blankState;
    setSheetNames((current) => [...current, name]);
    setActiveSheet(name);
    setState(blankState);
    setSelection({ anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } });
    setExtraSelections([]);
    setEditRequest(null);
    setFocusedCell(null);
    setDisplayedRowCount(GRID_ROW_TARGET);
    setDisplayedColumnCount(30);
    setLoading(false);
    showToast(`${name} 已创建`);
  }, [activeSheet, sheetNames, showToast, state]);

  const removeWorksheet = useCallback((sheetName: string) => {
    if (sheetNames.length <= 1) {
      showToast("至少保留一个工作表");
      return;
    }
    if (!window.confirm(`确定删除工作表“${sheetName}”吗？此操作不可撤销。`)) return;

    const index = sheetNames.indexOf(sheetName);
    const nextNames = sheetNames.filter((name) => name !== sheetName);
    delete sheetStatesRef.current[sheetName];
    setSheetNames(nextNames);

    if (activeSheet === sheetName) {
      const nextActive = nextNames[Math.max(0, Math.min(index, nextNames.length - 1))];
      const nextState = sheetStatesRef.current[nextActive] ?? createBlankSheetState();
      sheetStatesRef.current[nextActive] = nextState;
      setActiveSheet(nextActive);
      setState(nextState);
      setSelection({ anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } });
      setExtraSelections([]);
      setEditRequest(null);
      setFocusedCell(null);
    }
    showToast(`工作表“${sheetName}”已删除`);
  }, [activeSheet, sheetNames, showToast, state]);

  const deleteWorksheet = useCallback((sheetName: string, x = sheetContextPointRef.current.x, y = sheetContextPointRef.current.y) => {
    setSheetContextMenu({ sheet: sheetName, x, y });
  }, []);

  const renameWorksheet = useCallback((sheetName: string, value: string) => {
    setEditingSheetName(null);
    const nextName = value.trim();
    if (!nextName || nextName === sheetName) return;
    if (sheetName === PRIMARY_SHEET_NAME) {
      showToast("默认工作表不可重命名");
      return;
    }
    if (sheetNames.some((name) => name !== sheetName && name === nextName)) {
      showToast("工作表名称已存在");
      return;
    }
    const sheetState = sheetStatesRef.current[sheetName];
    if (sheetState) {
      delete sheetStatesRef.current[sheetName];
      sheetStatesRef.current[nextName] = sheetState;
    }
    setSheetNames((current) => current.map((name) => name === sheetName ? nextName : name));
    if (activeSheet === sheetName) setActiveSheet(nextName);
    showToast(`工作表“${sheetName}”已重命名为“${nextName}”`);
  }, [activeSheet, sheetNames, showToast]);

  useEffect(() => {
    const handleSheetDoubleClick = (event: MouseEvent) => {
      const target = (event.target as HTMLElement | null)?.closest(".sheet-tab") as HTMLElement | null;
      const sheetName = target?.textContent?.trim();
      if (!target || !sheetName) return;
      event.preventDefault();
      const nextName = window.prompt("重命名工作表", sheetName);
      if (nextName !== null) renameWorksheet(sheetName, nextName);
    };
    document.addEventListener("dblclick", handleSheetDoubleClick);
    return () => document.removeEventListener("dblclick", handleSheetDoubleClick);
  }, [renameWorksheet]);

  useEffect(() => {
    if (activeSheet === PRIMARY_SHEET_NAME) void refresh(INITIAL_STATE_LIMIT);
  }, [activeSheet, refresh]);

  useEffect(() => {
    if (activeSheet !== PRIMARY_SHEET_NAME || !identity.name || loading || state.rows.length >= GRID_ROW_TARGET) return;
    if (gridEnsureInFlightRef.current >= GRID_ROW_TARGET) return;
    gridEnsureInFlightRef.current = GRID_ROW_TARGET;
    void ensureGrid(identity, GRID_ROW_TARGET)
      .then(() => refresh())
      .catch(() => undefined)
      .finally(() => {
        if (gridEnsureInFlightRef.current === GRID_ROW_TARGET) gridEnsureInFlightRef.current = 0;
      });
  }, [activeSheet, identity, loading, refresh, state.rows.length]);

  const handleGridScroll = useCallback((position: { top: number; left: number; scrollHeight: number; scrollWidth: number; clientHeight: number; clientWidth: number }) => {
    if (position.scrollHeight - position.top - position.clientHeight < 850 && displayedRowCount < GRID_ROW_MAX) {
      // Keep extending in sizeable blocks; never stop at a cosmetic viewport
      // threshold while the server already has more rows available.
      const nextCount = Math.min(GRID_ROW_MAX, Math.max(displayedRowCount + 500, state.rows.length + 500));
      setDisplayedRowCount(nextCount);
      if (activeSheet !== PRIMARY_SHEET_NAME && state.rows.length < nextCount) {
        setState((current) => {
          const rows = [...current.rows];
          for (let id = current.next_row_id; rows.length < nextCount; id += 1) rows.push(createBlankRow(id));
          return { ...current, rows, next_row_id: rows.length + 1 };
        });
      } else if (activeSheet === PRIMARY_SHEET_NAME && state.rows.length < nextCount && gridEnsureInFlightRef.current < nextCount) {
        loadedRowLimitRef.current = nextCount;
        gridEnsureInFlightRef.current = nextCount;
        void ensureGrid(identity, nextCount).then(() => refresh(nextCount)).catch(() => undefined).finally(() => {
          if (gridEnsureInFlightRef.current === nextCount) gridEnsureInFlightRef.current = 0;
        });
      }
    }
    // Keep at least one more viewport of editable columns ready. Expanding by
    // a full viewport prevents a blank right side when a fast horizontal drag
    // lands beyond the columns that were previously rendered.
    const visibleWidth = Math.max(position.clientWidth, 1200);
    if (position.scrollWidth - position.left - position.clientWidth < visibleWidth && displayedColumnCount < FIELDS.length) {
      const averageWidth = Math.max(96, columnWidths.slice(0, Math.max(1, displayedColumnCount)).reduce((sum, width) => sum + width, 0) / Math.max(1, displayedColumnCount));
      const extraColumns = Math.max(16, Math.ceil((visibleWidth * 1.5) / averageWidth));
      setDisplayedColumnCount((current) => Math.min(FIELDS.length, current + extraColumns));
    }
  }, [activeSheet, columnWidths, displayedColumnCount, displayedRowCount, identity, refresh, state.rows.length]);

  useEffect(() => {
    if (!identity.name) return;
    setConnection("connecting");
    const source = new EventSource(eventsUrl(identity, loadedRowLimitRef.current));

    source.onopen = () => setConnection("live");
    source.onerror = () => setConnection("reconnecting");
    source.addEventListener("snapshot", (event) => {
      if (activeSheetRef.current !== PRIMARY_SHEET_NAME) return;
      const snapshot = JSON.parse((event as MessageEvent).data) as GridBookState;
      // A snapshot at the same revision can still be stale relative to an
      // optimistic local edit (for example, the initial SSE snapshot arriving
      // just after a border action). Only a strictly newer revision may replace
      // the current layout wholesale; the mutation response applies its own
      // authoritative layout below.
      setState((current) => {
        if (snapshot.revision <= current.revision && current.rows.length > 0) return current;
        const pendingFormats = Object.values(pendingFormatOperationsRef.current)
          .flatMap((operation) => Object.entries(operation));
        if (!pendingFormats.length) {
          sheetStatesRef.current[PRIMARY_SHEET_NAME] = snapshot;
          return snapshot;
        }
        const formats = { ...snapshot.layout.formats };
        for (const [key, format] of pendingFormats) {
          formats[key] = { ...(formats[key] ?? {}), ...format };
        }
        const merged = { ...snapshot, layout: { ...snapshot.layout, formats } };
        sheetStatesRef.current[PRIMARY_SHEET_NAME] = merged;
        return merged;
      });
      setConnection("live");
      setLoading(false);
    });
    source.addEventListener("presence", (event) => {
      if (activeSheetRef.current !== PRIMARY_SHEET_NAME) return;
      const payload = JSON.parse((event as MessageEvent).data) as { presence: PresenceItem[] };
      setState((current) => ({ ...current, presence: payload.presence }));
    });
    source.addEventListener("cell", (event) => {
      if (activeSheetRef.current !== PRIMARY_SHEET_NAME) return;
      const payload = JSON.parse((event as MessageEvent).data) as CellEvent;
      const key = `${payload.row_id}:${payload.field}`;
      setState((current) => ({
        ...current,
        revision: Math.max(current.revision, payload.revision),
        rows: current.rows.map((row) =>
          row.id === payload.row_id ? { ...row, [payload.field]: payload.value } : row,
        ),
        cells: { ...current.cells, [key]: payload.meta },
        activity: [payload.activity, ...current.activity.filter((item) => item.id !== payload.activity.id)],
      }));
      setFlashes((current) => new Set(current).add(key));
      window.setTimeout(() => {
        setFlashes((current) => {
          const next = new Set(current);
          next.delete(key);
          return next;
        });
      }, 1350);
    });
    source.addEventListener("row", () => { if (activeSheetRef.current === PRIMARY_SHEET_NAME) void refresh(); });
    source.addEventListener("grid", () => { if (activeSheetRef.current === PRIMARY_SHEET_NAME) void refresh(); });
    source.addEventListener("rows_deleted", () => { if (activeSheetRef.current === PRIMARY_SHEET_NAME) void refresh(); });
    source.addEventListener("layout", (event) => {
      if (activeSheetRef.current !== PRIMARY_SHEET_NAME) return;
      const payload = JSON.parse((event as MessageEvent).data) as LayoutEvent;
      setState((current) => ({
        ...current,
        revision: Math.max(current.revision, payload.revision),
        layout: payload.revision > current.revision
          ? (() => {
              const pendingFormats = Object.values(pendingFormatOperationsRef.current)
                .flatMap((operation) => Object.entries(operation));
              if (!pendingFormats.length) return payload.layout;
              const formats = { ...payload.layout.formats };
              for (const [key, format] of pendingFormats) {
                formats[key] = { ...(formats[key] ?? {}), ...format };
              }
              return { ...payload.layout, formats };
            })()
          : current.layout,
      }));
    });

    return () => source.close();
  }, [identity, refresh]);

  useEffect(() => {
    setColumnWidths(
      FIELDS.map((field) => state.layout.column_widths[field.key]
        ?? DEFAULT_WORKSHEET_COLUMN_WIDTH),
    );
  }, [state.layout.column_widths]);

  useEffect(() => {
    if (!identity.name) return;
    const send = () => void updatePresence(identity, focusedCell).catch(() => undefined);
    send();
    const interval = window.setInterval(send, 20_000);
    return () => window.clearInterval(interval);
  }, [focusedCell, identity]);

  useEffect(() => () => {
    if (toastTimer.current) window.clearTimeout(toastTimer.current);
  }, []);

  useEffect(() => {
    const finishSelection = () => {
      // Commit the last pointer target before cancelling the queued frame. A
      // fast mouseup can arrive between the final mouseenter and its rAF;
      // flushing here keeps the selection endpoint exact without forcing
      // every pointer event through React.
      const target = selectionDragTarget.current;
      if (target) {
        setSelection((current) => (
          current.focus.row === target.row && current.focus.col === target.col
            ? current
            : { ...current, focus: target }
        ));
      }
      draggingSelection.current = false;
      selectionDragTarget.current = null;
      if (selectionDragFrame.current !== null) {
        window.cancelAnimationFrame(selectionDragFrame.current);
        selectionDragFrame.current = null;
      }
      formulaReferenceDragRef.current = null;
    };
    window.addEventListener("mouseup", finishSelection);
    window.addEventListener("blur", finishSelection);
    return () => {
      window.removeEventListener("mouseup", finishSelection);
      window.removeEventListener("blur", finishSelection);
    };
  }, []);

  const commitCell = useCallback(
    async (rowId: number, field: EditableField, value: EditableCellValue, version: number) => {
      const previousRow = state.rows.find((row) => row.id === rowId);
      const previousValue = previousRow?.[field];
      const before: EditableCellValue = previousValue === undefined ? null : previousValue as EditableCellValue;
      if (before !== value) {
        historyUndoStack.current.push({
          kind: "value",
          item: { updates: [{ row_id: rowId, field, before, after: value }], label: "编辑单元格" },
        });
        historyRedoStack.current = [];
      }
      setState((current) => ({
        ...current,
        rows: current.rows.map((row) => row.id === rowId ? { ...row, [field]: value } : row),
        cells: { ...current.cells, [`${rowId}:${field}`]: { ...(current.cells[`${rowId}:${field}`] ?? {}), version, updated_by: identity.name, client_id: identity.clientId, updated_at: new Date().toISOString() } },
      }));
      if (activeSheet !== PRIMARY_SHEET_NAME) return;
      try {
        await updateCell(identity, rowId, field, value, version);
      } catch (error) {
        if (error instanceof ApiError && error.status === 409) {
          showToast("这格刚被其他人更新，已为你载入最新内容");
          await refresh();
        } else {
          showToast(error instanceof Error ? error.message : "保存失败，请重试");
        }
        throw error;
      }
    },
    [activeSheet, identity, refresh, showToast, state.rows],
  );

  const commitFormulaDraft = useCallback(async () => {
    if (!currentCell) return;
    await commitCell(currentCell.row.id, currentCell.field.key, formulaDraft, state.cells[`${currentCell.row.id}:${currentCell.field.key}`]?.version ?? 0);
  }, [commitCell, currentCell, formulaDraft, state.cells]);

  useEffect(() => {
    const handle: FormulaEditorHandle = {
      beginReference: (cell) => {
        const input = formulaInputRef.current;
        const draft = formulaDraftRef.current;
        const cursorRange = formulaCursorRange(input, draft);
        const cursorStart = cursorRange.start;
        const cursorEnd = cursorRange.end;
        const prefix = draft.slice(0, cursorStart);
        const suffix = draft.slice(cursorEnd);
        formulaBarReferenceBaseRef.current = { prefix, suffix };
        const reference = formulaReference(cell);
        setFormulaDraft(`${prefix}${reference}${suffix}`);
        requestAnimationFrame(() => {
          formulaInputRef.current?.focus();
          const nextCursor = prefix.length + reference.length;
          formulaInputRef.current?.setSelectionRange(nextCursor, nextCursor);
        });
      },
      updateReference: (start, end) => {
        const base = formulaBarReferenceBaseRef.current;
        if (!base) return;
        const reference = formulaReference(start, end);
        setFormulaDraft(`${base.prefix}${reference}${base.suffix}`);
        requestAnimationFrame(() => {
          formulaInputRef.current?.focus();
          const nextCursor = base.prefix.length + reference.length;
          formulaInputRef.current?.setSelectionRange(nextCursor, nextCursor);
        });
      },
    };
    formulaBarEditorRef.current = handle;
    return () => {
      if (formulaBarEditorRef.current === handle) formulaBarEditorRef.current = null;
    };
  }, []);

  const registerFormulaEditor = useCallback((handle: FormulaEditorHandle | null) => {
    cellFormulaEditorRef.current = handle;
  }, []);

  const activeFormulaEditor = useCallback(() => {
    const formulaBarFocused = formulaInputRef.current === document.activeElement
      && formulaDraftRef.current.trimStart().startsWith("=");
    return formulaBarFocused ? formulaBarEditorRef.current : cellFormulaEditorRef.current;
  }, []);

  const formulaReferencePointerDown = useCallback((row: number, col: number) => {
    const editor = activeFormulaEditor();
    if (!editor) return false;
    const start = { row, col };
    formulaReferenceDragRef.current = { editor, start };
    editor.beginReference(start);
    return true;
  }, [activeFormulaEditor]);

  const formulaReferencePointerMove = useCallback((row: number, col: number) => {
    const session = formulaReferenceDragRef.current;
    if (!session) return false;
    session.editor.updateReference(session.start, { row, col });
    return true;
  }, []);

  const chooseFormulaInBar = useCallback((name: string) => {
    const input = formulaInputRef.current;
    const draft = formulaDraftRef.current;
    const cursor = input?.selectionStart ?? draft.length;
    const before = draft.slice(0, cursor);
    const token = before.match(/[A-Za-z]*$/)?.[0] ?? "";
    const prefix = before.slice(0, before.length - token.length);
    const suffix = draft.slice(input?.selectionEnd ?? cursor);
    const next = `${prefix}${name}()${suffix}`;
    setFormulaDraft(next);
    requestAnimationFrame(() => {
      formulaInputRef.current?.focus();
      const nextCursor = prefix.length + name.length + 1;
      formulaInputRef.current?.setSelectionRange(nextCursor, nextCursor);
    });
  }, []);

  const selectCell = useCallback((row: number, col: number, extend: boolean, additive = false) => {
    setSelection((current) => ({
      anchor: extend ? current.anchor : { row, col },
      focus: { row, col },
    }));
    if (!extend) setExtraSelections((current) => additive ? [...current, selection] : []);
    setContextMenu(null);
    tableCardRef.current?.focus({ preventScroll: true });
  }, [selection]);

  const selectMergedRange = useCallback((row: number, col: number, rowSpan: number, colSpan: number, extend: boolean, additive = false) => {
    setSelection((current) => ({
      anchor: extend ? current.anchor : { row, col },
      focus: { row: row + rowSpan - 1, col: col + colSpan - 1 },
    }));
    if (!extend) setExtraSelections((current) => additive ? [...current, selection] : []);
    setContextMenu(null);
    tableCardRef.current?.focus({ preventScroll: true });
  }, [selection]);

  const beginCellSelection = useCallback((row: number, col: number, extend: boolean, additive = false) => {
    draggingSelection.current = true;
    selectCell(row, col, extend, additive);
  }, [selectCell]);

  const dragSelectionTo = useCallback((row: number, col: number) => {
    if (!draggingSelection.current) return;
    selectionDragTarget.current = { row, col };
    if (selectionDragFrame.current !== null) return;
    selectionDragFrame.current = window.requestAnimationFrame(() => {
      selectionDragFrame.current = null;
      const target = selectionDragTarget.current;
      if (!target || !draggingSelection.current) return;
      setSelection((current) => (
        current.focus.row === target.row && current.focus.col === target.col
          ? current
          : { ...current, focus: target }
      ));
    });
  }, []);

  const moveActiveCell = useCallback((rowDelta: number, columnDelta: number) => {
    setExtraSelections([]);
    setSelection((current) => {
      const row = Math.max(0, Math.min(state.rows.length - 1, current.focus.row + rowDelta));
      const col = Math.max(0, Math.min(FIELDS.length - 1, current.focus.col + columnDelta));
      return { anchor: { row, col }, focus: { row, col } };
    });
    tableCardRef.current?.focus({ preventScroll: true });
  }, [state.rows.length]);

  const requestCellEdit = useCallback((row: number, col: number, seed: string | null) => {
    editRequestId.current += 1;
    setEditRequest({ row, col, seed, id: editRequestId.current });
  }, []);

  const selectRow = useCallback((row: number, extend: boolean, additive = false) => {
    draggingSelection.current = true;
    setSelection((current) => ({
      anchor: extend ? current.anchor : { row, col: 0 },
      focus: { row, col: FIELDS.length - 1 },
    }));
    if (!extend) setExtraSelections((current) => additive ? [...current, selection] : []);
    tableCardRef.current?.focus({ preventScroll: true });
  }, [selection]);

  const selectColumn = useCallback((col: number, extend: boolean, additive = false) => {
    draggingSelection.current = true;
    const bottom = Math.max(0, state.rows.length - 1);
    setSelection((current) => ({
      anchor: extend ? current.anchor : { row: 0, col },
      focus: { row: bottom, col },
    }));
    if (!extend) setExtraSelections((current) => additive ? [...current, selection] : []);
    tableCardRef.current?.focus({ preventScroll: true });
  }, [selection, state.rows.length]);

  const dragRowSelection = useCallback((row: number) => {
    if (!draggingSelection.current) return;
    dragSelectionTo(row, FIELDS.length - 1);
  }, [dragSelectionTo]);

  const dragColumnSelection = useCallback((col: number) => {
    if (!draggingSelection.current) return;
    dragSelectionTo(Math.max(0, state.rows.length - 1), col);
  }, [dragSelectionTo, state.rows.length]);

  const selectAll = useCallback(() => {
    if (!state.rows.length) return;
    setExtraSelections([]);
    setSelection({
      anchor: { row: 0, col: 0 },
      focus: { row: state.rows.length - 1, col: FIELDS.length - 1 },
    });
    tableCardRef.current?.focus({ preventScroll: true });
  }, [state.rows.length]);

  const beginColumnResize = useCallback((columnIndex: number, event: ReactMouseEvent) => {
    event.preventDefault();
    event.stopPropagation();
    const startX = event.clientX;
    const startWidth = columnWidths[columnIndex];
    let finalWidth = startWidth;
    const onMove = (moveEvent: MouseEvent) => {
      const nextWidth = Math.max(72, Math.min(420, startWidth + moveEvent.clientX - startX));
      finalWidth = Math.round(nextWidth);
      setColumnWidths((current) => current.map((width, index) => index === columnIndex ? nextWidth : width));
    };
    const onUp = () => {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
      if (activeSheet === PRIMARY_SHEET_NAME) {
        void updateLayout(identity, "column_width", {
          field: FIELDS[columnIndex].key,
          width: finalWidth,
        }).catch((error) => showToast(error instanceof Error ? error.message : "列宽保存失败"));
      } else {
        setState((current) => ({
          ...current,
          layout: { ...current.layout, column_widths: { ...current.layout.column_widths, [FIELDS[columnIndex].key]: finalWidth } },
        }));
      }
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
  }, [activeSheet, columnWidths, identity, showToast]);

  const beginRowResize = useCallback((row: GridBookRow, event: ReactMouseEvent) => {
    event.preventDefault();
    event.stopPropagation();
    const startY = event.clientY;
    const savedHeight = state.layout.row_heights[String(row.id)];
    const startHeight = savedHeight && savedHeight !== LEGACY_DEFAULT_ROW_HEIGHT
      ? savedHeight
      : DEFAULT_WORKSHEET_ROW_HEIGHT;
    let finalHeight = startHeight;
    const onMove = (moveEvent: MouseEvent) => {
      finalHeight = Math.round(Math.max(30, Math.min(240, startHeight + moveEvent.clientY - startY)));
      setState((current) => ({
        ...current,
        layout: {
          ...current.layout,
          row_heights: { ...current.layout.row_heights, [String(row.id)]: finalHeight },
        },
      }));
    };
    const onUp = () => {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
      if (activeSheet === PRIMARY_SHEET_NAME) {
        void updateLayout(identity, "row_height", { row_id: row.id, height: finalHeight })
          .catch((error) => showToast(error instanceof Error ? error.message : "行高保存失败"));
      }
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
  }, [activeSheet, identity, showToast, state.layout.row_heights]);

  const allSelections = useMemo(() => [...extraSelections, selection], [extraSelections, selection]);
  const hasRangeSelection = allSelections.length > 1 || allSelections.some((item) => {
    const bounds = selectionBounds(item);
    return bounds.top !== bounds.bottom || bounds.left !== bounds.right;
  });
  const selectionName = useMemo(() => selectionLabel(allSelections), [allSelections]);

  const isCellSelected = useCallback((row: number, col: number) => allSelections.some((item) => {
    const bounds = selectionBounds(item);
    return row >= bounds.top && row <= bounds.bottom && col >= bounds.left && col <= bounds.right;
  }), [allSelections]);

  const selectedValues = useCallback(() => {
    const output: unknown[][] = [];
    allSelections.forEach((item, rangeIndex) => {
      if (rangeIndex) output.push([]);
      const bounds = selectionBounds(item);
      for (let rowIndex = bounds.top; rowIndex <= bounds.bottom; rowIndex += 1) {
        const row = state.rows[rowIndex];
        if (!row) continue;
        const values: unknown[] = [];
        for (let columnIndex = bounds.left; columnIndex <= bounds.right; columnIndex += 1) {
          const field = FIELDS[columnIndex];
          if (field) values.push(row[field.key]);
        }
        output.push(values);
      }
    });
    return output;
  }, [allSelections, state.rows]);

  const selectedCoordinates = useCallback(() => {
    const cells: Array<{ row_id: number; field: EditableField }> = [];
    const seen = new Set<string>();
    allSelections.forEach((item) => {
      const bounds = selectionBounds(item);
      for (let rowIndex = bounds.top; rowIndex <= bounds.bottom; rowIndex += 1) {
        const row = state.rows[rowIndex];
        if (!row) continue;
        for (let columnIndex = bounds.left; columnIndex <= bounds.right; columnIndex += 1) {
          const field = FIELDS[columnIndex];
          const key = field ? `${row.id}:${field.key}` : "";
          if (field && !seen.has(key)) { seen.add(key); cells.push({ row_id: row.id, field: field.key }); }
        }
      }
    });
    return cells;
  }, [allSelections, state.rows]);

  const copySelection = useCallback(async () => {
    const text = serializeSpreadsheetClipboard(selectedValues());
    try {
      await navigator.clipboard.writeText(text);
      showToast("已复制");
    } catch {
      showToast("浏览器未允许剪贴板，请按 Ctrl+C");
    }
    setContextMenu(null);
  }, [selectedValues, showToast]);

  const submitBatch = useCallback(async (
    updates: Array<{
      row_id: number;
      field: EditableField;
      value: string | number | null;
      version: number;
    }>,
    message: string,
  ) => {
    if (!updates.length) return;
    const rowById = new Map(state.rows.map((row) => [row.id, row]));
    const historyUpdates = updates
      .map((update) => {
        const previousValue = rowById.get(update.row_id)?.[update.field];
        const before: EditableCellValue = previousValue === undefined ? null : previousValue as EditableCellValue;
        return { row_id: update.row_id, field: update.field, before, after: update.value };
      })
      .filter((update) => update.before !== update.after);
    if (historyUpdates.length) {
      historyUndoStack.current.push({ kind: "value", item: { updates: historyUpdates, label: message } });
      historyRedoStack.current = [];
    }
    setPasting(true);
    try {
      if (activeSheet !== PRIMARY_SHEET_NAME) {
        setState((current) => {
          const byRow = new Map<number, Array<{ field: EditableField; value: string | number | null }>>();
          updates.forEach((update) => byRow.set(update.row_id, [...(byRow.get(update.row_id) ?? []), update]));
          const rows = current.rows.map((row) => {
            const rowUpdates = byRow.get(row.id);
            if (!rowUpdates) return row;
            const next = { ...row } as GridBookRow;
            rowUpdates.forEach((update) => { next[update.field] = update.value as never; });
            return next;
          });
          return { ...current, rows };
        });
        showToast(message);
        return;
      }
      try {
        await submitOperations(
          identity,
          state.revision,
          updates.map((update) => ({ type: "cell", ...update })),
        );
      } catch (error) {
        // Only old demo servers lack the operation endpoint. Never replay an
        // operation after a conflict or an unknown network/server failure.
        if (error instanceof ApiError && error.status === 404) {
          await updateCells(identity, updates);
        } else {
          throw error;
        }
      }
      showToast(message);
    } catch (error) {
      if (error instanceof ApiError && error.status === 409) {
        showToast("粘贴区域刚被其他人更新，已载入最新内容");
        await refresh();
      } else {
        showToast(error instanceof Error ? error.message : "批量更新失败，请重试");
      }
      throw error;
    } finally {
      setPasting(false);
    }
  }, [activeSheet, identity, refresh, showToast, state.revision, state.rows]);

  const submitFillBatch = useCallback(async (
    updates: Array<{
      row_id: number;
      field: EditableField;
      value: string | number | null;
      version: number;
    }>,
    formats: Array<{ row_id: number; field: EditableField; format: CellFormat }>,
    message: string,
  ) => {
    if (!updates.length) return;
    const rowById = new Map(state.rows.map((row) => [row.id, row]));
    const historyUpdates = updates
      .map((update) => ({
        row_id: update.row_id,
        field: update.field,
        before: (rowById.get(update.row_id)?.[update.field] ?? null) as EditableCellValue,
        after: update.value,
      }))
      .filter((update) => update.before !== update.after);
    if (historyUpdates.length) {
      historyUndoStack.current.push({ kind: "value", item: { updates: historyUpdates, label: message } });
      historyRedoStack.current = [];
    }
    setPasting(true);
    try {
      if (activeSheet !== PRIMARY_SHEET_NAME) {
        showToast(message);
        return;
      }
      try {
        await submitOperations(identity, state.revision, [{ type: "fill", updates, formats }]);
      } catch (error) {
        if (!(error instanceof ApiError) || error.status !== 404) throw error;
        await submitOperations(identity, state.revision, updates.map((update) => ({ type: "cell", ...update })));
        const groups = new Map<string, { format: CellFormat; cells: Array<{ row_id: number; field: EditableField }> }>();
        formats.forEach((item) => {
          const key = JSON.stringify(item.format);
          const group = groups.get(key) ?? { format: item.format, cells: [] };
          group.cells.push({ row_id: item.row_id, field: item.field });
          groups.set(key, group);
        });
        for (const group of groups.values()) {
          await updateLayout(identity, "format", { cells: group.cells, format: group.format, replace: true });
        }
      }
      showToast(message);
    } catch (error) {
      if (error instanceof ApiError && error.status === 409) {
        showToast("填充区域刚被其他人更新，已载入最新内容");
        await refresh();
      } else {
        showToast(error instanceof Error ? error.message : "填充失败，请重试");
      }
      throw error;
    } finally {
      setPasting(false);
    }
  }, [activeSheet, identity, refresh, showToast, state.revision, state.rows]);

  const fillSelection = useCallback(async (source: CellSelection, target: CellPosition, copyMode = false, doubleClick = false) => {
    const sourceBounds = selectionBounds(source);
    const targetBounds: FillBounds = {
      top: Math.max(0, Math.min(sourceBounds.top, target.row)),
      bottom: Math.min(state.rows.length - 1, Math.max(sourceBounds.bottom, target.row)),
      left: Math.max(0, Math.min(sourceBounds.left, target.col)),
      right: Math.min(FIELDS.length - 1, Math.max(sourceBounds.right, target.col)),
    };
    if (doubleClick) {
      const adjacentColumns = [sourceBounds.left - 1, sourceBounds.right + 1].filter((column) => column >= 0 && column < FIELDS.length);
      const adjacentColumn = adjacentColumns.find((column) => {
        const value = state.rows[sourceBounds.top]?.[FIELDS[column]?.key];
        return value !== null && value !== undefined && String(value).trim() !== "";
      });
      if (adjacentColumn === undefined) return;
      let lastRow = sourceBounds.bottom;
      while (lastRow + 1 < state.rows.length) {
        const value = state.rows[lastRow + 1]?.[FIELDS[adjacentColumn]?.key];
        if (value === null || value === undefined || String(value).trim() === "") break;
        lastRow += 1;
      }
      targetBounds.bottom = lastRow;
    }
    const grid = { rows: state.rows, fields: FIELDS, formats: state.layout.formats };
    const updates: Array<{ row_id: number; field: EditableField; value: string | number | null; version: number }> = [];
    const formatUpdates: Array<{ row_id: number; field: EditableField; format: CellFormat }> = [];

    for (let rowIndex = targetBounds.top; rowIndex <= targetBounds.bottom; rowIndex += 1) {
      const row = state.rows[rowIndex];
      if (!row) continue;
      for (let columnIndex = targetBounds.left; columnIndex <= targetBounds.right; columnIndex += 1) {
        const inSource = rowIndex >= sourceBounds.top && rowIndex <= sourceBounds.bottom && columnIndex >= sourceBounds.left && columnIndex <= sourceBounds.right;
        if (inSource) continue;
        const field = FIELDS[columnIndex];
        if (!field) continue;
        const generated = fillCellAt(grid, sourceBounds, rowIndex, columnIndex, copyMode ? "copy" : "auto");
        updates.push({ row_id: row.id, field: field.key, value: generated.value, version: state.cells[`${row.id}:${field.key}`]?.version ?? 0 });
        formatUpdates.push({ row_id: row.id, field: field.key, format: generated.format });
      }
    }
    if (!updates.length) return;

    // Apply the result immediately so dragging never waits on the network.
    setState((current) => {
      const byRow = new Map<number, Array<{ field: EditableField; value: string | number | null }>>();
      updates.forEach((update) => byRow.set(update.row_id, [...(byRow.get(update.row_id) ?? []), update]));
      const nextRows = current.rows.map((row) => {
        const rowUpdates = byRow.get(row.id);
        if (!rowUpdates) return row;
        const next = { ...row } as GridBookRow;
        rowUpdates.forEach((update) => { next[update.field] = update.value as never; });
        return next;
      });
      const formats = { ...current.layout.formats };
      formatUpdates.forEach((item) => { formats[`${item.row_id}:${item.field}`] = { ...item.format }; });
      return { ...current, rows: nextRows, layout: { ...current.layout, formats } };
    });
    setExtraSelections([]);
    setSelection({ anchor: { row: targetBounds.top, col: targetBounds.left }, focus: { row: targetBounds.bottom, col: targetBounds.right } });
    try {
      await submitFillBatch(updates, formatUpdates, `已填充 ${updates.length} 个单元格`);
    } catch {
      // submitBatch already reports conflicts and refreshes the latest state.
    }
  }, [showToast, state.cells, state.layout.formats, state.rows, submitFillBatch]);

  const pasteSpreadsheetText = useCallback(async (text: string) => {
    const matrix = parseSpreadsheetClipboard(text);
    const bounds = selectionBounds(selection);
    const targetRows = bounds.bottom - bounds.top + 1;
    const targetColumns = bounds.right - bounds.left + 1;
    const fillSelection = matrix.length === 1 && matrix[0].length === 1 && (targetRows > 1 || targetColumns > 1);
    const rowCount = fillSelection ? targetRows : matrix.length;
    const columnCount = fillSelection ? targetColumns : Math.max(...matrix.map((row) => row.length));
    const updates: Array<{
      row_id: number;
      field: EditableField;
      value: string | number | null;
      version: number;
    }> = [];
    let clipped = false;

    for (let rowOffset = 0; rowOffset < rowCount; rowOffset += 1) {
      const rowIndex = bounds.top + rowOffset;
      const row = state.rows[rowIndex];
      if (!row) {
        clipped = true;
        continue;
      }
      for (let columnOffset = 0; columnOffset < columnCount; columnOffset += 1) {
        const columnIndex = bounds.left + columnOffset;
        const field = FIELDS[columnIndex];
        if (!field) {
          clipped = true;
          continue;
        }
        const raw = fillSelection ? matrix[0][0] : (matrix[rowOffset]?.[columnOffset] ?? "");
        const value: string | number | null = raw;
        const cellKey = `${row.id}:${field.key}`;
        updates.push({
          row_id: row.id,
          field: field.key,
          value,
          version: state.cells[cellKey]?.version ?? 0,
        });
      }
    }

    if (!updates.length) {
      showToast("没有可粘贴的单元格");
      return;
    }
    try {
      await submitBatch(
        updates,
        clipped ? `已粘贴 ${updates.length} 格，超出表格的内容已忽略` : `已粘贴 ${updates.length} 个单元格`,
      );
      const lastRow = Math.min(state.rows.length - 1, bounds.top + rowCount - 1);
      const lastColumn = Math.min(FIELDS.length - 1, bounds.left + columnCount - 1);
      setSelection({ anchor: { row: bounds.top, col: bounds.left }, focus: { row: lastRow, col: lastColumn } });
    } catch {
      // Error feedback and state refresh are handled by submitBatch.
    }
  }, [selection, showToast, state.cells, state.rows, submitBatch]);

  const pasteFromSystemClipboard = useCallback(async () => {
    setContextMenu(null);
    try {
      if (navigator.clipboard.read) {
        const items = await navigator.clipboard.read();
        for (const item of items) {
          if (item.types.includes("text/html")) {
            const html = await (await item.getType("text/html")).text();
            const matrix = parseHtmlSpreadsheetClipboard(html);
            if (matrix.length) {
              await pasteSpreadsheetText(serializeSpreadsheetClipboard(matrix));
              return;
            }
          }
        }
      }
      const text = await navigator.clipboard.readText();
      if (!text) {
        showToast("剪贴板中没有可粘贴的表格内容");
        return;
      }
      await pasteSpreadsheetText(text);
    } catch {
      showToast("浏览器未允许读取剪贴板，请直接按 Ctrl+V");
    }
  }, [pasteSpreadsheetText, showToast]);

  const clearSelection = useCallback(async () => {
    setContextMenu(null);
    const bounds = selectionBounds(selection);
    const updates: Array<{
      row_id: number;
      field: EditableField;
      value: string | number | null;
      version: number;
    }> = [];
    for (let rowIndex = bounds.top; rowIndex <= bounds.bottom; rowIndex += 1) {
      const row = state.rows[rowIndex];
      if (!row) continue;
      for (let columnIndex = bounds.left; columnIndex <= bounds.right; columnIndex += 1) {
        const field = FIELDS[columnIndex];
        if (!field) continue;
        updates.push({
          row_id: row.id,
          field: field.key,
          value: "",
          version: state.cells[`${row.id}:${field.key}`]?.version ?? 0,
        });
      }
    }
    if (!updates.length) return;
    try {
      await submitBatch(updates, `已清空 ${updates.length} 个单元格`);
    } catch {
      // Error feedback is handled by submitBatch.
    }
  }, [selection, showToast, state.cells, state.rows, submitBatch]);

  const cutSelection = useCallback(async () => {
    await copySelection();
    await clearSelection();
  }, [clearSelection, copySelection]);

  const applySelectedFormat = useCallback(async (format: CellFormat) => {
    const cells = selectedCoordinates();
    if (!cells.length) return;
    const before = cells.map((cell) => state.layout.formats[`${cell.row_id}:${cell.field}`] ?? {});
    historyUndoStack.current.push({ kind: "format", item: { cells, before, after: format } });
    historyRedoStack.current = [];
    setState((current) => {
      const formats = { ...current.layout.formats };
      for (const cell of cells) {
        const key = `${cell.row_id}:${cell.field}`;
        formats[key] = { ...(formats[key] ?? {}), ...format };
      }
      return { ...current, layout: { ...current.layout, formats } };
    });
    try {
      if (activeSheet === PRIMARY_SHEET_NAME) await formatCells(identity, cells, format);
    } catch (error) {
      showToast(error instanceof Error ? error.message : "格式设置失败");
    }
  }, [activeSheet, identity, selectedCoordinates, showToast, state.layout.formats]);

  const applyBorderAction = useCallback(async (action: BorderAction) => {
    const cells = selectedCoordinates();
    if (!cells.length) return;
    const bounds = selectionBounds(selection);
    const edge = { style: "solid" as const, width: 1 as const, color: "#66717d" };
    const operationId = ++formatOperationIdRef.current;
    const pendingFormats: Record<string, CellFormat> = {};
    for (const cell of cells) {
      const rowIndex = state.rows.findIndex((row) => row.id === cell.row_id);
      const colIndex = FIELDS.findIndex((field) => field.key === cell.field);
      const key = `${cell.row_id}:${cell.field}`;
      const next = { ...(state.layout.formats[key] ?? {}) };
      const borders: CellBorders = { ...(next.borders ?? {}) };
      if (next.border && !next.borders) {
        borders.top = edge; borders.right = edge; borders.bottom = edge; borders.left = edge;
      }
      if (action === "none") {
        next.borders = {};
        next.border = false;
      } else {
        const sides: Array<keyof CellBorders> = action === "all"
          ? ["top", "right", "bottom", "left"]
          : action === "outside"
            ? [
                ...(rowIndex === bounds.top ? ["top" as const] : []),
                ...(rowIndex === bounds.bottom ? ["bottom" as const] : []),
                ...(colIndex === bounds.left ? ["left" as const] : []),
                ...(colIndex === bounds.right ? ["right" as const] : []),
              ]
            : [action as keyof CellBorders];
        sides.forEach((side) => { borders[side] = edge; });
        next.borders = borders;
        next.border = true;
      }
      pendingFormats[key] = next;
    }
    pendingFormatOperationsRef.current[operationId] = pendingFormats;
    setState((current) => {
      const formats = { ...current.layout.formats };
      for (const cell of cells) {
        const rowIndex = current.rows.findIndex((row) => row.id === cell.row_id);
        const colIndex = FIELDS.findIndex((field) => field.key === cell.field);
        const key = `${cell.row_id}:${cell.field}`;
        const next = { ...(formats[key] ?? {}) };
        const borders: CellBorders = { ...(next.borders ?? {}) };
        if (next.border && !next.borders) {
          borders.top = edge; borders.right = edge; borders.bottom = edge; borders.left = edge;
        }
        if (action === "none") {
          next.borders = {};
          next.border = false;
        } else {
          const sides: Array<keyof CellBorders> = action === "all"
            ? ["top", "right", "bottom", "left"]
            : action === "outside"
              ? [
                  ...(rowIndex === bounds.top ? ["top" as const] : []),
                  ...(rowIndex === bounds.bottom ? ["bottom" as const] : []),
                  ...(colIndex === bounds.left ? ["left" as const] : []),
                  ...(colIndex === bounds.right ? ["right" as const] : []),
                ]
              : [action as keyof CellBorders];
          sides.forEach((side) => { borders[side] = edge; });
          next.borders = borders;
          next.border = true;
        }
        formats[key] = next;
      }
      return { ...current, layout: { ...current.layout, formats } };
    });
    setBorderMenuOpen(false);
    try {
      if (activeSheet === PRIMARY_SHEET_NAME) {
        const response = await updateBorders(identity, cells, action, edge);
        const revision = typeof response.revision === "number" ? response.revision : Number(response.revision);
        if (response.layout && Number.isFinite(revision)) {
          setState((current) => {
            // A cell/presence event from another client can advance the
            // revision before this format response arrives.  Treating that
            // response as stale used to drop the optimistic border and make
            // the border appear to roll back.  Use the response layout only
            // when it is current; otherwise keep the latest local layout and
            // overlay the still-pending format operations.
      const responseLayout = response.layout as GridBookState["layout"];
            const layout = revision >= current.revision ? responseLayout : current.layout;
            const pendingFormats = Object.values(pendingFormatOperationsRef.current)
              .flatMap((operation) => Object.entries(operation));
            if (!pendingFormats.length) {
              return {
                ...current,
                revision: Math.max(current.revision, revision),
                layout,
              };
            }
            const formats = { ...layout.formats };
            for (const [key, format] of pendingFormats) {
              formats[key] = { ...(formats[key] ?? {}), ...format };
            }
            return {
              ...current,
              revision: Math.max(current.revision, revision),
              layout: { ...layout, formats },
            };
          });
        }
      }
    } catch (error) {
      showToast(error instanceof Error ? error.message : "框线设置失败");
    } finally {
      delete pendingFormatOperationsRef.current[operationId];
    }
  }, [activeSheet, identity, selectedCoordinates, selection, showToast, state]);

  const undoFormatLegacy = useCallback(async () => {
    const item = formatUndoStack.current.pop();
    if (!item) { showToast("没有可撤销的操作"); return; }
    setState((current) => {
      const formats = { ...current.layout.formats };
      item.cells.forEach((cell, index) => { formats[`${cell.row_id}:${cell.field}`] = item.before[index] ?? {}; });
      return { ...current, layout: { ...current.layout, formats } };
    });
    formatRedoStack.current.push(item);
    if (activeSheet === PRIMARY_SHEET_NAME) await Promise.all(item.cells.map((cell, index) => formatCells(identity, [cell], item.before[index] ?? {}))).catch(() => undefined);
  }, [activeSheet, identity, showToast]);

  const redoFormatLegacy = useCallback(async () => {
    const item = formatRedoStack.current.pop();
    if (!item) { showToast("没有可重做的操作"); return; }
    setState((current) => {
      const formats = { ...current.layout.formats };
      item.cells.forEach((cell) => { formats[`${cell.row_id}:${cell.field}`] = { ...(formats[`${cell.row_id}:${cell.field}`] ?? {}), ...item.after }; });
      return { ...current, layout: { ...current.layout, formats } };
    });
    formatUndoStack.current.push(item);
    if (activeSheet === PRIMARY_SHEET_NAME) await Promise.all(item.cells.map((cell) => formatCells(identity, [cell], item.after))).catch(() => undefined);
  }, [activeSheet, identity, showToast]);

  const applyValueHistory = useCallback(async (item: ValueHistoryItem, direction: "undo" | "redo") => {
    const updates = item.updates.map((update) => ({
      row_id: update.row_id,
      field: update.field,
      value: direction === "undo" ? update.before : update.after,
      version: state.cells[`${update.row_id}:${update.field}`]?.version ?? 0,
    }));
    setState((current) => {
      const byRow = new Map<number, typeof updates>();
      updates.forEach((update) => byRow.set(update.row_id, [...(byRow.get(update.row_id) ?? []), update]));
      const rows = current.rows.map((row) => {
        const rowUpdates = byRow.get(row.id);
        if (!rowUpdates) return row;
        const next = { ...row } as GridBookRow;
        rowUpdates.forEach((update) => { next[update.field] = update.value as never; });
        return next;
      });
      return { ...current, rows };
    });
    if (activeSheet !== PRIMARY_SHEET_NAME) return;
    try {
      try {
        await submitOperations(identity, state.revision, updates.map((update) => ({ type: "cell", ...update })));
      } catch (error) {
        if (error instanceof ApiError && error.status === 404) await updateCells(identity, updates);
        else throw error;
      }
    } catch (error) {
      showToast(error instanceof Error ? error.message : "撤销或重做失败");
      throw error;
    }
  }, [activeSheet, identity, showToast, state.cells, state.revision]);

  const applyFormatHistory = useCallback(async (item: FormatHistoryItem, direction: "undo" | "redo") => {
    setState((current) => {
      const formats = { ...current.layout.formats };
      item.cells.forEach((cell, index) => {
        const key = `${cell.row_id}:${cell.field}`;
        formats[key] = direction === "undo"
          ? item.before[index] ?? {}
          : { ...(formats[key] ?? {}), ...item.after };
      });
      return { ...current, layout: { ...current.layout, formats } };
    });
    if (activeSheet !== PRIMARY_SHEET_NAME) return;
    try {
      if (direction === "undo") {
        await Promise.all(item.cells.map((cell, index) => formatCells(identity, [cell], item.before[index] ?? {})));
      } else {
        await Promise.all(item.cells.map((cell) => formatCells(identity, [cell], item.after)));
      }
    } catch (error) {
      showToast(error instanceof Error ? error.message : "格式撤销或重做失败");
      throw error;
    }
  }, [activeSheet, identity, showToast]);

  const undoFormat = useCallback(async () => {
    const action = historyUndoStack.current.pop();
    if (!action) { showToast("没有可撤销的操作"); return; }
    try {
      if (action.kind === "value") await applyValueHistory(action.item, "undo");
      else await applyFormatHistory(action.item, "undo");
      historyRedoStack.current.push(action);
    } catch {
      historyUndoStack.current.push(action);
    }
  }, [applyFormatHistory, applyValueHistory, showToast]);

  const redoFormat = useCallback(async () => {
    const action = historyRedoStack.current.pop();
    if (!action) { showToast("没有可重做的操作"); return; }
    try {
      if (action.kind === "value") await applyValueHistory(action.item, "redo");
      else await applyFormatHistory(action.item, "redo");
      historyUndoStack.current.push(action);
    } catch {
      historyRedoStack.current.push(action);
    }
  }, [applyFormatHistory, applyValueHistory, showToast]);

  useEffect(() => {
    if (!formatBrush) return undefined;
    const applyBrush = () => {
      const target = `${selection.focus.row}:${selection.focus.col}`;
      if (target === formatBrushSource.current) return;
      void applySelectedFormat(formatBrush);
      showToast("格式已应用");
      if (!formatBrushContinuous) setFormatBrush(null);
    };
    window.addEventListener("mouseup", applyBrush);
    return () => window.removeEventListener("mouseup", applyBrush);
  }, [applySelectedFormat, formatBrush, formatBrushContinuous, selection.focus.col, selection.focus.row, showToast]);

  const mergesIntersectingSelection = useCallback(() => {
    const bounds = selectionBounds(selection);
    const rowPositions = new Map(state.rows.map((row, index) => [row.id, index]));
    return state.layout.merges.filter((merge) => {
      const top = rowPositions.get(merge.start_row_id);
      const bottom = rowPositions.get(merge.end_row_id);
      const left = FIELDS.findIndex((field) => field.key === merge.start_field);
      const right = FIELDS.findIndex((field) => field.key === merge.end_field);
      if (top === undefined || bottom === undefined || left < 0 || right < 0) return false;
      return !(bounds.bottom < top || bounds.top > bottom || bounds.right < left || bounds.left > right);
    });
  }, [selection, state.layout.merges, state.rows]);

  const mergeSelection = useCallback(async () => {
    const bounds = selectionBounds(selection);
    const startRow = state.rows[bounds.top];
    const endRow = state.rows[bounds.bottom];
    if (!startRow || !endRow) return;
    try {
      if (activeSheet === PRIMARY_SHEET_NAME) {
        await updateLayout(identity, "merge", {
          start_row_id: startRow.id,
          end_row_id: endRow.id,
          start_field: FIELDS[bounds.left].key,
          end_field: FIELDS[bounds.right].key,
        });
      }
      setState((current) => ({
        ...current,
        layout: {
          ...current.layout,
          merges: [...current.layout.merges, { id: `local-${Date.now()}`, start_row_id: startRow.id, end_row_id: endRow.id, start_field: FIELDS[bounds.left].key, end_field: FIELDS[bounds.right].key }],
        },
      }));
      setSelection({
        anchor: { row: bounds.bottom, col: bounds.right },
        focus: { row: bounds.top, col: bounds.left },
      });
      showToast("已合并选中单元格；其他格的数据仍被安全保留");
    } catch (error) {
      showToast(error instanceof Error ? error.message : "合并失败");
    }
  }, [activeSheet, identity, selection, showToast, state.rows]);

  const unmergeSelection = useCallback(async () => {
    const mergeIds = mergesIntersectingSelection().map((merge) => merge.id);
    if (!mergeIds.length) {
      showToast("选区中没有合并单元格");
      return;
    }
    try {
      if (activeSheet === PRIMARY_SHEET_NAME) await updateLayout(identity, "unmerge", { merge_ids: mergeIds });
      setState((current) => ({
        ...current,
        layout: { ...current.layout, merges: current.layout.merges.filter((merge) => !mergeIds.includes(merge.id)) },
      }));
      showToast("已取消合并");
    } catch (error) {
      showToast(error instanceof Error ? error.message : "取消合并失败");
    }
  }, [activeSheet, identity, mergesIntersectingSelection, showToast]);

  const insertSelectedRow = useCallback(async (position: "above" | "below") => {
    const bounds = selectionBounds(selection);
    const anchorIndex = position === "above" ? bounds.top : bounds.bottom;
    const afterIndex = position === "above" ? anchorIndex - 1 : anchorIndex;
    const afterRowId = afterIndex >= 0 ? state.rows[afterIndex]?.id : 0;
    try {
      if (activeSheet !== PRIMARY_SHEET_NAME) {
        const insertAt = position === "above" ? bounds.top : bounds.bottom + 1;
        setState((current) => {
          const rows = [...current.rows];
          rows.splice(insertAt, 0, createBlankRow(current.next_row_id));
          return { ...current, rows, next_row_id: current.next_row_id + 1 };
        });
      } else {
        await addRow(identity, afterRowId);
      }
      showToast(position === "above" ? "已在选区上方插入一行" : "已在选区下方插入一行");
    } catch (error) {
      showToast(error instanceof Error ? error.message : "插入行失败");
    }
  }, [activeSheet, identity, selection, showToast, state.rows]);

  const deleteSelectedRows = useCallback(async () => {
    const bounds = selectionBounds(selection);
    const rowIds = state.rows.slice(bounds.top, bounds.bottom + 1).map((row) => row.id);
    try {
      if (activeSheet !== PRIMARY_SHEET_NAME) {
        setState((current) => {
          const rows = current.rows.filter((_, index) => index < bounds.top || index > bounds.bottom);
          const nextRows = rows.length ? rows : [createBlankRow(current.next_row_id)];
          return { ...current, rows: nextRows, next_row_id: Math.max(current.next_row_id, nextRows[nextRows.length - 1].id + 1) };
        });
      } else {
        await deleteRows(identity, rowIds);
      }
      showToast(`已删除 ${rowIds.length} 行`);
      setSelection({ anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } });
    } catch (error) {
      showToast(error instanceof Error ? error.message : "删除行失败");
    }
  }, [activeSheet, identity, selection, showToast, state.rows]);

  const resetSelectedDimensions = useCallback(async (kind: "row" | "column") => {
    const bounds = selectionBounds(selection);
    try {
      if (kind === "column") {
        const fields = FIELDS.slice(bounds.left, bounds.right + 1);
        if (activeSheet === PRIMARY_SHEET_NAME) {
          await Promise.all(fields.map((field) => updateLayout(identity, "column_width", { field: field.key, width: DEFAULT_WORKSHEET_COLUMN_WIDTH })));
        } else {
          setColumnWidths((current) => current.map((width, index) => index >= bounds.left && index <= bounds.right ? DEFAULT_WORKSHEET_COLUMN_WIDTH : width));
          setState((current) => ({
            ...current,
            layout: {
              ...current.layout,
              column_widths: {
                ...current.layout.column_widths,
                ...Object.fromEntries(fields.map((field) => [field.key, DEFAULT_WORKSHEET_COLUMN_WIDTH])),
              },
            },
          }));
        }
        showToast("已恢复标准列宽");
      } else {
        const rows = state.rows.slice(bounds.top, bounds.bottom + 1);
        if (activeSheet === PRIMARY_SHEET_NAME) {
          await Promise.all(rows.map((row) => updateLayout(identity, "row_height", { row_id: row.id, height: DEFAULT_WORKSHEET_ROW_HEIGHT })));
        } else {
          setState((current) => ({
            ...current,
            layout: { ...current.layout, row_heights: Object.fromEntries(Object.entries(current.layout.row_heights).filter(([key]) => !rows.some((row) => String(row.id) === key))) },
          }));
        }
        showToast("已恢复标准行高");
      }
    } catch (error) {
      showToast(error instanceof Error ? error.message : "尺寸设置失败");
    }
  }, [activeSheet, identity, selection, showToast, state.rows]);

  const handleCopyEvent = (event: ClipboardEvent<HTMLElement>) => {
    if ((event.target as HTMLElement).closest("input, textarea, [contenteditable='true']")) return;
    event.preventDefault();
    event.clipboardData.setData("text/plain", serializeSpreadsheetClipboard(selectedValues()));
    showToast("已复制");
  };

  const handleCutEvent = (event: ClipboardEvent<HTMLElement>) => {
    if ((event.target as HTMLElement).closest("input, textarea, [contenteditable='true']")) return;
    event.preventDefault();
    event.clipboardData.setData("text/plain", serializeSpreadsheetClipboard(selectedValues()));
    void clearSelection();
  };

  const handlePasteEvent = (event: ClipboardEvent<HTMLElement>) => {
    if ((event.target as HTMLElement).closest("input, textarea, [contenteditable='true']")) return;
    event.preventDefault();
    if (pasting) return;
    const html = event.clipboardData.getData("text/html");
    const matrix = html ? parseHtmlSpreadsheetClipboard(html) : [];
    void pasteSpreadsheetText(matrix.length ? serializeSpreadsheetClipboard(matrix) : event.clipboardData.getData("text/plain"));
  };

  const handleTableKeyDown = (event: KeyboardEvent<HTMLElement>) => {
    if ((event.target as HTMLElement).closest("input, textarea, [contenteditable='true']")) return;
    const shortcut = (event.ctrlKey || event.metaKey) && !event.altKey;
    const shortcutKey = event.key.toLowerCase();
    const zoomInKey = event.key === "+" || event.key === "=" || event.code === "Equal";
    const zoomOutKey = event.key === "-" || event.key === "_" || event.code === "Minus";
    if (event.key === "Escape") {
      event.preventDefault();
      setContextMenu(null);
      setFormatPalette(null);
      setFormatPanelOpen(false);
      setMenuOpen(false);
      setEditRequest(null);
      return;
    }
    if (shortcut && zoomInKey) {
      event.preventDefault();
      setZoom((current) => Math.min(200, current + 10));
      return;
    }
    if (shortcut && zoomOutKey) {
      event.preventDefault();
      setZoom((current) => Math.max(50, current - 10));
      return;
    }
    if (shortcut && shortcutKey === "0") {
      event.preventDefault();
      setZoom(100);
      return;
    }
    if (shortcut && shortcutKey === "s") {
      event.preventDefault();
      showToast("已自动保存");
      return;
    }
    if (shortcut && shortcutKey === "c") {
      event.preventDefault();
      void copySelection();
      return;
    }
    if (shortcut && shortcutKey === "x") {
      event.preventDefault();
      void cutSelection();
      return;
    }
    if (shortcut && shortcutKey === "v") {
      event.preventDefault();
      void pasteFromSystemClipboard();
      return;
    }
    if (shortcut && shortcutKey === "b") {
      event.preventDefault();
      void applySelectedFormat({ bold: !activeFormat.bold });
      return;
    }
    if (shortcut && shortcutKey === "i") {
      event.preventDefault();
      void applySelectedFormat({ italic: !activeFormat.italic });
      return;
    }
    if (shortcut && shortcutKey === "u") {
      event.preventDefault();
      void applySelectedFormat({ underline: !activeFormat.underline });
      return;
    }
    if (shortcut && shortcutKey === "z") {
      event.preventDefault();
      if (event.shiftKey) void redoFormat(); else void undoFormat();
      return;
    }
    if (shortcut && shortcutKey === "y") {
      event.preventDefault();
      void redoFormat();
      return;
    }
    if (shortcut && shortcutKey === "f") {
      event.preventDefault();
      setFindPanelOpen(true);
      return;
    }
    if (shortcut && shortcutKey === "h") {
      event.preventDefault();
      setFindPanelOpen(true);
      return;
    }
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "a") {
      event.preventDefault();
      selectAll();
      return;
    }
    if (event.key === "Home") {
      event.preventDefault();
      const row = shortcut ? 0 : selection.focus.row;
      selectCell(row, 0, event.shiftKey);
      return;
    }
    if (event.key === "End") {
      event.preventDefault();
      const row = shortcut ? Math.max(0, state.rows.length - 1) : selection.focus.row;
      const col = Math.max(0, FIELDS.length - 1);
      selectCell(row, col, event.shiftKey);
      return;
    }
    if (shortcut && event.key === " ") {
      event.preventDefault();
      const col = selection.focus.col;
      setSelection({ anchor: { row: 0, col }, focus: { row: Math.max(0, state.rows.length - 1), col } });
      return;
    }
    if (event.shiftKey && event.key === " ") {
      event.preventDefault();
      const row = selection.focus.row;
      setSelection({ anchor: { row, col: 0 }, focus: { row, col: Math.max(0, FIELDS.length - 1) } });
      return;
    }
    if (shortcut && (event.key === "PageDown" || event.key === "PageUp")) {
      event.preventDefault();
      const currentIndex = sheetNames.indexOf(activeSheet);
      const direction = event.key === "PageDown" ? 1 : -1;
      const nextIndex = Math.max(0, Math.min(sheetNames.length - 1, currentIndex + direction));
      if (nextIndex !== currentIndex) void switchSheet(sheetNames[nextIndex]);
      return;
    }
    if (event.key === "PageDown" || event.key === "PageUp") {
      event.preventDefault();
      const delta = event.key === "PageDown" ? 15 : -15;
      const row = Math.max(0, Math.min(state.rows.length - 1, selection.focus.row + delta));
      selectCell(row, selection.focus.col, false);
      return;
    }
    if (shortcut && event.key.startsWith("Arrow")) {
      event.preventDefault();
      const row = event.key === "ArrowUp" || event.key === "ArrowDown"
        ? (event.key === "ArrowUp" ? 0 : Math.max(0, state.rows.length - 1))
        : selection.focus.row;
      const col = event.key === "ArrowLeft" || event.key === "ArrowRight"
        ? (event.key === "ArrowLeft" ? 0 : Math.max(0, FIELDS.length - 1))
        : selection.focus.col;
      selectCell(row, col, event.shiftKey);
      return;
    }
    const keyMoves: Record<string, [number, number]> = {
      ArrowUp: [-1, 0],
      ArrowDown: [1, 0],
      ArrowLeft: [0, -1],
      ArrowRight: [0, 1],
    };
    const move = keyMoves[event.key];
    if (move) {
      event.preventDefault();
      const row = Math.max(0, Math.min(state.rows.length - 1, selection.focus.row + move[0]));
      const col = Math.max(0, Math.min(FIELDS.length - 1, selection.focus.col + move[1]));
      selectCell(row, col, event.shiftKey);
    } else if (event.key === "Tab") {
      event.preventDefault();
      moveActiveCell(0, event.shiftKey ? -1 : 1);
    } else if (event.key === "Enter" || event.key === "F2") {
      event.preventDefault();
      requestCellEdit(selection.focus.row, selection.focus.col, null);
    } else if (
      event.key.length === 1 &&
      !event.ctrlKey &&
      !event.metaKey &&
      !event.altKey
    ) {
      event.preventDefault();
      requestCellEdit(selection.focus.row, selection.focus.col, event.key);
    } else if (event.key === "Delete" || event.key === "Backspace") {
      event.preventDefault();
      void clearSelection();
    }
  };

  const handleAddRow = async () => {
    setAddingRow(true);
    try {
      if (activeSheet !== PRIMARY_SHEET_NAME) {
        setState((current) => ({
          ...current,
          rows: [...current.rows, createBlankRow(current.next_row_id)],
          next_row_id: current.next_row_id + 1,
        }));
      } else {
        await addRow(identity);
      }
      showToast("已新增一行，可直接填写任意单元格");
    } catch (error) {
      showToast(error instanceof Error ? error.message : "新增失败，请重试");
    } finally {
      setAddingRow(false);
    }
  };

  const exportCsv = () => {
    const escape = (value: unknown) => {
      const text = value === null || value === undefined ? "" : String(value);
      return /[",\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
    };
    const lines = [
      FIELDS.map((field) => escape(field.label)).join(","),
      ...state.rows.map((row) => FIELDS.map((field) => escape(row[field.key])).join(",")),
    ];
    const blob = new Blob(["\ufeff", lines.join("\r\n")], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `GridBook-${state.period.end || "export"}.csv`;
    anchor.click();
    URL.revokeObjectURL(url);
    showToast("CSV 已导出");
  };

  const saveIdentity = (name: string) => {
    localStorage.setItem(NAME_KEY, name);
    setIdentity((current) => ({ ...current, name }));
    setIdentityOpen(false);
  };

  const onlineUsers = useMemo(
    () => state.presence.filter((item, index, items) => items.findIndex((other) => other.client_id === item.client_id) === index),
    [state.presence],
  );
  const selectedBounds = useMemo(() => selectionBounds(selection), [selection]);
  const selectedCellCount = selectedCoordinates().length;
  const selectionStatistics = useMemo(() => {
    const formulaContext = { rows: state.rows, fields: FIELDS };
    const values = selectedValues()
      .flat()
      .map((value) => displayFormulaValue(value, formulaContext))
      .filter((value) => value !== "" && value !== null && value !== undefined);
    const numbers = values.map((value) => Number(value)).filter((value) => Number.isFinite(value));
    const sum = numbers.reduce((total, value) => total + value, 0);
    return {
      count: values.length,
      sum,
      average: numbers.length ? sum / numbers.length : 0,
    };
  }, [selectedValues, state.rows]);
  const activeRow = state.rows[selection.focus.row];
  const activeField = FIELDS[selection.focus.col];
  const activeFormat: CellFormat =
    activeRow && activeField
      ? state.layout.formats[`${activeRow.id}:${activeField.key}`] ?? {}
      : {};
  useEffect(() => {
    const handleGlobalShortcut = (event: globalThis.KeyboardEvent) => {
      const target = event.target as HTMLElement | null;
      if (event.defaultPrevented) return;
      if (target?.closest("input, textarea, [contenteditable='true']")) return;
      if (event.key === "Escape") {
        event.preventDefault();
        setContextMenu(null);
        setFormatPalette(null);
        setFormatPanelOpen(false);
        setMenuOpen(false);
        setEditRequest(null);
        return;
      }
      if (!(event.ctrlKey || event.metaKey) || event.altKey) return;
      const key = event.key.toLowerCase();
      if (key === "s") {
        event.preventDefault();
        showToast("已自动保存");
        return;
      }
      const inGrid = Boolean(target?.closest(".excel-table-card, .virtual-grid-shell, .grid-cell, .grid-row-head, .grid-col-head"));
      if (!inGrid) return;
      const zoomInKey = event.key === "+" || event.key === "=" || event.code === "Equal";
      const zoomOutKey = event.key === "-" || event.key === "_" || event.code === "Minus";
      if (zoomInKey) { event.preventDefault(); setZoom((current) => Math.min(200, current + 10)); }
      else if (zoomOutKey) { event.preventDefault(); setZoom((current) => Math.max(50, current - 10)); }
      else if (key === "0") { event.preventDefault(); setZoom(100); }
      else if (event.key.startsWith("Arrow")) {
        event.preventDefault();
        const row = event.key === "ArrowUp" || event.key === "ArrowDown"
          ? (event.key === "ArrowUp" ? 0 : Math.max(0, state.rows.length - 1))
          : selection.focus.row;
        const col = event.key === "ArrowLeft" || event.key === "ArrowRight"
          ? (event.key === "ArrowLeft" ? 0 : Math.max(0, FIELDS.length - 1))
          : selection.focus.col;
        selectCell(row, col, event.shiftKey);
      }
      else if (key === "c") { event.preventDefault(); void copySelection(); }
      else if (key === "x") { event.preventDefault(); void cutSelection(); }
      else if (key === "v") { event.preventDefault(); void pasteFromSystemClipboard(); }
      else if (key === "a") { event.preventDefault(); selectAll(); }
      else if (key === "b") { event.preventDefault(); void applySelectedFormat({ bold: !activeFormat.bold }); }
      else if (key === "i") { event.preventDefault(); void applySelectedFormat({ italic: !activeFormat.italic }); }
      else if (key === "u") { event.preventDefault(); void applySelectedFormat({ underline: !activeFormat.underline }); }
      else if (key === "z") { event.preventDefault(); if (event.shiftKey) void redoFormat(); else void undoFormat(); }
      else if (key === "y") { event.preventDefault(); void redoFormat(); }
      else if (key === "f" || key === "h") { event.preventDefault(); setFindPanelOpen(true); }
    };
    window.addEventListener("keydown", handleGlobalShortcut);
    return () => window.removeEventListener("keydown", handleGlobalShortcut);
  }, [activeFormat.bold, activeFormat.italic, activeFormat.underline, applySelectedFormat, copySelection, cutSelection, pasteFromSystemClipboard, redoFormat, selectAll, selection.focus.col, selection.focus.row, selectCell, showToast, state.rows.length, undoFormat]);
  useEffect(() => {
    const handleGridZoomWheel = (event: WheelEvent) => {
      if (!(event.ctrlKey || event.metaKey) || event.defaultPrevented) return;
      const target = event.target as HTMLElement | null;
      if (!target?.closest(".excel-table-card, .virtual-grid-shell")) return;
      event.preventDefault();
      event.stopPropagation();
      const direction = event.deltaY > 0 ? -1 : 1;
      setZoom((current) => Math.max(50, Math.min(200, current + direction * 10)));
    };
    document.addEventListener("wheel", handleGridZoomWheel, { capture: true, passive: false });
    return () => document.removeEventListener("wheel", handleGridZoomWheel, { capture: true });
  }, []);
  const selectedMerges = mergesIntersectingSelection();
  const mergeGeometry = useMemo(() => {
    const hidden = new Set<string>();
    const anchors = new Map<string, { rowSpan: number; colSpan: number; id: string }>();
    const rowPositions = new Map(state.rows.map((row, index) => [row.id, index]));
    for (const merge of state.layout.merges) {
      const top = rowPositions.get(merge.start_row_id);
      const bottom = rowPositions.get(merge.end_row_id);
      const left = FIELDS.findIndex((field) => field.key === merge.start_field);
      const right = FIELDS.findIndex((field) => field.key === merge.end_field);
      if (top === undefined || bottom === undefined || left < 0 || right < 0) continue;
      anchors.set(`${top}:${left}`, {
        rowSpan: bottom - top + 1,
        colSpan: right - left + 1,
        id: merge.id,
      });
      for (let row = top; row <= bottom; row += 1) {
        for (let col = left; col <= right; col += 1) {
          if (row !== top || col !== left) hidden.add(`${row}:${col}`);
        }
      }
    }
    return { hidden, anchors };
  }, [state.layout.merges, state.rows]);
  const latestUpdate = state.activity[0]?.updated_at;
  const displayedRows = useMemo(
    () => state.rows.slice(0, Math.min(Math.max(displayedRowCount, state.rows.length), GRID_ROW_MAX)),
    [displayedRowCount, state.rows],
  );
  const findMatches = useMemo(() => {
    const matches: Array<{ row: number; col: number }> = [];
    const normalized = findQuery.trim().toLocaleLowerCase();
    if (!normalized) return matches;
    state.rows.forEach((row, rowIndex) => {
      FIELDS.forEach((field, col) => {
        if (String(row[field.key] ?? "").toLocaleLowerCase().includes(normalized)) {
          matches.push({ row: rowIndex, col });
        }
      });
    });
    return matches;
  }, [findQuery, state.rows]);
  const findMatchKeys = useMemo(
    () => new Set(findMatches.map((match) => `${match.row}:${match.col}`)),
    [findMatches],
  );
  const moveToFindMatch = useCallback((direction: 1 | -1) => {
    if (!findMatches.length) {
      showToast(findQuery.trim() ? "未找到匹配内容" : "请输入要查找的内容");
      return;
    }
    const next = (findCursor + direction + findMatches.length) % findMatches.length;
    const match = findMatches[next];
    setFindCursor(next);
    setSelection({ anchor: match, focus: match });
    tableCardRef.current?.focus({ preventScroll: true });
  }, [findCursor, findMatches, findQuery, showToast]);
  const replaceCurrentMatch = useCallback(async () => {
    const match = findMatches[findCursor];
    if (!match || !findQuery.trim()) {
      showToast("请先输入要替换的内容");
      return;
    }
    const row = state.rows[match.row];
    const field = FIELDS[match.col];
    if (!row || !field) return;
    const current = String(row[field.key] ?? "");
    const next = current.replace(findQuery, replaceText);
    if (next === current) {
      moveToFindMatch(1);
      return;
    }
    try {
      await submitBatch([{ row_id: row.id, field: field.key, value: next, version: state.cells[`${row.id}:${field.key}`]?.version ?? 0 }], "已替换当前单元格");
      moveToFindMatch(1);
    } catch {
      // submitBatch displays the failure reason.
    }
  }, [findCursor, findMatches, findQuery, moveToFindMatch, replaceText, showToast, state.cells, state.rows, submitBatch]);
  const replaceAllMatches = useCallback(async () => {
    if (!findQuery.trim()) {
      showToast("请先输入要替换的内容");
      return;
    }
    const updates = findMatches.flatMap(({ row: rowIndex, col }) => {
      const row = state.rows[rowIndex];
      const field = FIELDS[col];
      if (!row || !field) return [];
      const current = String(row[field.key] ?? "");
      const value = current.replaceAll(findQuery, replaceText);
      return value === current ? [] : [{ row_id: row.id, field: field.key, value, version: state.cells[`${row.id}:${field.key}`]?.version ?? 0 }];
    });
    try {
      await submitBatch(updates, `已替换 ${updates.length} 个单元格`);
      setFindCursor(0);
    } catch {
      // submitBatch displays the failure reason.
    }
  }, [findMatches, findQuery, replaceText, showToast, state.cells, state.rows, submitBatch]);
  const resetSelectedFormat = useCallback(() => {
    void applySelectedFormat({
      bold: false,
      italic: false,
      underline: false,
      border: false,
      wrap: false,
      align: "left",
      vertical_align: "middle",
      fill: "none",
      font_color: "default",
      font_size: 12,
      number_format: "general",
    });
    setFormatPanelOpen(false);
    showToast("已恢复默认单元格格式");
  }, [applySelectedFormat, showToast]);
  const copyShareLink = useCallback(async () => {
    try {
      await navigator.clipboard.writeText(window.location.href);
      showToast("分享链接已复制");
    } catch {
      window.prompt("复制此链接", window.location.href);
    }
  }, [showToast]);

  return (
    <div className="app-shell excel-app-shell" onContextMenu={(event) => event.preventDefault()} onMouseDown={() => setContextMenu(null)}>
      <header className="topbar">
        <div className="brand-block">
          <div className="brand-mark" aria-hidden="true"><span /><span /><span /><span /></div>
          <h1>空白表格</h1>
          <span className="sheet-save-state">✓ 所有编辑内容都会自动保存到云端</span>
        </div>
        <div className="topbar-right">
          <button className="sheet-icon-button" onClick={() => setMenuOpen((current) => !current)} type="button" aria-label="菜单">☰</button>
          <button className="sheet-icon-button" onClick={() => setIdentityOpen(true)} type="button" aria-label="协作">♙</button>
          <button className="sheet-share-button" onClick={() => void copyShareLink()} type="button">分享</button>
          <button className="identity-button" onClick={() => setIdentityOpen(true)} type="button"><Avatar clientId={identity.clientId} name={identity.name || "用户"} /></button>
        </div>
      </header>
      <main className="workspace excel-workspace">
        <section className="workspace-main">
          <div className="spreadsheet-toolbar excel-toolbar" aria-label="表格工具栏">
            <div className="toolbar-group">
              <button onClick={() => void undoFormat()} aria-label="撤销" type="button"><Undo2 /></button>
              <button onClick={() => void redoFormat()} aria-label="重做" type="button"><Redo2 /></button>
              <button className={formatBrush ? "active" : ""} onClick={(event) => {
                if (formatBrush && event.detail >= 2) { setFormatBrushContinuous(true); showToast("连续格式刷已开启"); return; }
                if (formatBrush && formatBrushContinuous) { setFormatBrush(null); setFormatBrushContinuous(false); showToast("已退出连续格式刷"); return; }
                if (!formatBrush) { formatBrushSource.current = `${selection.focus.row}:${selection.focus.col}`; setFormatBrush(activeFormat); setFormatBrushContinuous(false); showToast("格式刷已就绪，请拖选目标区域"); return; }
                if (formatBrush) { setFormatBrush(null); showToast("已退出格式刷"); }
                else {
                  formatBrushSource.current = `${selection.focus.row}:${selection.focus.col}`;
                  setFormatBrush(activeFormat);
                  showToast("格式刷已就绪，请选择目标区域");
                }
              }} aria-label="格式刷" type="button"><Paintbrush />格式刷</button>
              <button onClick={() => void copySelection()} aria-label="复制" type="button"><Copy />复制</button>
              <button onClick={() => void cutSelection()} aria-label="剪切" type="button"><Scissors />剪切</button>
              <button onClick={() => void pasteFromSystemClipboard()} aria-label="粘贴" type="button"><ClipboardPaste />粘贴</button>
            </div>
            <div className="toolbar-group formatting-group">
              <ToolbarSelect className="font-family-select" ariaLabel="字体" value={activeFormat.font_family ?? "Microsoft YaHei"} options={FONT_OPTIONS} onChange={(value) => void applySelectedFormat({ font_family: value })} />
              <ToolbarSelect ariaLabel="字号" value={activeFormat.font_size ?? 12} options={[9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36].map((size) => ({ value: size, label: String(size) }))} onChange={(value) => void applySelectedFormat({ font_size: value })} />
              <button className={activeFormat.bold ? "active" : ""} onClick={() => void applySelectedFormat({ bold: !activeFormat.bold })} aria-label="加粗" type="button"><RoundedTextFormatIcon kind="bold" /></button>
              <button className={activeFormat.italic ? "active" : ""} onClick={() => void applySelectedFormat({ italic: !activeFormat.italic })} aria-label="斜体" type="button"><RoundedTextFormatIcon kind="italic" /></button>
              <button className={activeFormat.underline ? "active" : ""} onClick={() => void applySelectedFormat({ underline: !activeFormat.underline })} aria-label="下划线" type="button"><RoundedTextFormatIcon kind="underline" /></button>
              <span className="toolbar-menu-anchor border-menu-trigger">
                <button className={borderMenuOpen ? "active" : ""} onClick={() => { setBorderMenuOpen((current) => !current); setFormatPalette(null); setFormatPanelOpen(false); setFindPanelOpen(false); }} aria-label="框线" aria-expanded={borderMenuOpen} type="button"><Grid3X3 />框线</button>
                {borderMenuOpen ? <BorderMenu onSelect={(action) => void applyBorderAction(action)} /> : null}
              </span>
              <span className="toolbar-menu-anchor">
                <button className={formatPalette === "fill" ? "active" : ""} onClick={() => setFormatPalette((current) => current === "fill" ? null : "fill")} aria-label="填充" type="button"><PaintBucket />填充</button>
                {formatPalette === "fill" ? <ColorPalette title="填充" recentColors={RECENT_COLORS} allowNone onSelect={(color) => { void applySelectedFormat({ fill: color }); setFormatPalette(null); }} /> : null}
              </span>
              <span className="toolbar-menu-anchor">
                <button className={formatPalette === "font" ? "active" : ""} onClick={() => setFormatPalette((current) => current === "font" ? null : "font")} aria-label="字色" type="button"><Palette />字色</button>
                {formatPalette === "font" ? <ColorPalette title="字色" recentColors={RECENT_COLORS} onSelect={(color) => { void applySelectedFormat({ font_color: color }); setFormatPalette(null); }} /> : null}
              </span>
              <button onClick={() => void clearSelection()} aria-label="清除" type="button"><Eraser />清除</button>
            </div>
            <div className="toolbar-group formatting-group">
              <button onClick={() => void applySelectedFormat({ vertical_align: "top" })} aria-label="顶端对齐" type="button"><AlignStartVertical /></button>
              <button onClick={() => void applySelectedFormat({ vertical_align: "middle" })} aria-label="垂直居中" type="button"><AlignCenterVertical /></button>
              <button onClick={() => void applySelectedFormat({ vertical_align: "bottom" })} aria-label="底端对齐" type="button"><AlignEndVertical /></button>
              <button onClick={() => void applySelectedFormat({ wrap: !activeFormat.wrap })} aria-label="换行" type="button"><WrapText /></button>
              <button onClick={() => void applySelectedFormat({ align: "left" })} aria-label="左对齐" type="button"><AlignLeft /></button>
              <button onClick={() => void applySelectedFormat({ align: "center" })} aria-label="居中对齐" type="button"><AlignCenter /></button>
              <button onClick={() => void applySelectedFormat({ align: "right" })} aria-label="右对齐" type="button"><AlignRight /></button>
              <button onClick={() => void applySelectedFormat({ align: "justify" })} aria-label="两端对齐" type="button"><AlignJustify /></button>
            </div>
            <div className="toolbar-group">
              <button onClick={() => void (selectedMerges.length ? unmergeSelection() : mergeSelection())} aria-label={selectedMerges.length ? "拆分单元格" : "合并单元格"} type="button"><TableCellsMerge />合并</button>
              <button className={`format-panel-trigger${formatPanelOpen ? " active" : ""}`} onClick={(event) => { setFormatPanelPosition(getFlyoutPosition(event, 280, 360)); setFormatPalette(null); setFindPanelOpen(false); setFormatPanelOpen((current) => !current); }} aria-label="格式" type="button"><TableProperties />格式</button>
              <button className={freeze.row ? "active" : ""} onClick={() => setFreeze((current) => ({ row: !current.row, column: false }))} aria-label="冻结表头" title="冻结表头" type="button"><Snowflake /></button>
              <button className={findPanelOpen ? "active" : ""} onClick={(event) => { setFindPanelPosition(getFlyoutPosition(event, 720, 90)); setFormatPanelOpen(false); setFindPanelOpen(true); setFindCursor(0); }} aria-label="查找" type="button"><Search /></button>
            </div>
          </div>
          {formatPanelOpen ? <div className="format-panel" role="dialog" aria-label="单元格格式" style={{ top: formatPanelPosition.top, left: formatPanelPosition.left }}>
            <div><strong>单元格格式</strong></div>
            <div className="number-format-list">
              {[
                ["general", "常规", "不使用特定格式"],
                ["number", "数值", "12,345.67"],
                ["date", "日期", "2026/8/13"],
                ["percentage", "百分比", "12.50%"],
                ["text", "文本", "按输入内容显示"],
              ].map(([value, label, hint]) => <button className={activeFormat.number_format === value || (!activeFormat.number_format && value === "general") ? "active" : ""} key={value} onClick={() => { void applySelectedFormat({ number_format: value as NonNullable<CellFormat["number_format"]> }); setFormatPanelOpen(false); }} type="button"><b>{label}</b><span>{hint}</span></button>)}
            </div>
          </div> : null}
          {findPanelOpen ? <div className="find-panel" role="dialog" aria-label="查找和替换" style={{ top: findPanelPosition.top, left: findPanelPosition.left }}>
            <input autoFocus aria-label="查找" onChange={(event) => { setFindQuery(event.target.value); setFindCursor(0); }} placeholder="查找" value={findQuery} />
            <input aria-label="替换为" onChange={(event) => setReplaceText(event.target.value)} placeholder="替换为" value={replaceText} />
            {findQuery.trim() ? <span>{`${findMatches.length} 项匹配`}</span> : null}
            <button onClick={() => moveToFindMatch(-1)} type="button">上一个</button><button onClick={() => moveToFindMatch(1)} type="button">下一个</button><button onClick={() => void replaceCurrentMatch()} type="button">替换</button><button onClick={() => void replaceAllMatches()} type="button">全部替换</button>
            <button aria-label="关闭查找和替换" className="panel-close" onClick={() => setFindPanelOpen(false)} title="关闭" type="button"><X /></button>
          </div> : null}
          <div className="formula-bar" aria-label="内容编辑栏">
            <div className="formula-name-card">
              <input aria-label="名称框" className="formula-name-box" value={selectionName || currentCellName} readOnly />
            </div>
            <div className="formula-editor-card">
              <span className="formula-fx" aria-hidden="true">fx</span>
              {!hasRangeSelection ? <FormulaSuggestions value={formulaDraft} onChoose={chooseFormulaInBar} /> : null}
              <div className="formula-content-wrap" title={hasRangeSelection ? "多选区域" : (formulaDraft || "当前单元格内容")}>
                {!hasRangeSelection && formulaDraft.trimStart().startsWith("=") ? <span className="formula-reference-preview" aria-hidden="true"><FormulaReferencePreview value={formulaDraft} /></span> : null}
                <input aria-label="单元格内容" className={`formula-content-input${formulaDraft.trimStart().startsWith("=") ? " has-formula-references" : ""}`} onChange={(event) => setFormulaDraft(event.target.value)} onKeyDown={(event) => {
                  if (event.key === "Enter") { event.preventDefault(); void commitFormulaDraft(); }
                  if (event.key === "Escape") { event.preventDefault(); setFormulaDraft(currentCellValue); }
                }} onBlur={() => setFormulaBarFocused(false)} onFocus={() => { setFormulaDraft(currentCellValue); setFormulaBarFocused(true); }} ref={formulaInputRef} value={hasRangeSelection ? "" : (formulaBarFocused ? formulaDraft : currentCellValue)} readOnly={hasRangeSelection} aria-disabled={hasRangeSelection} />
                {!hasRangeSelection && formulaDraft.length > 28 ? <span className="formula-full-value-hint" role="status">{formulaDraft}</span> : null}
              </div>
            </div>
          </div>
          <div className="sheet-tabs" aria-label="工作表标签">
            <button aria-label="新增工作表" className="sheet-add-button" onClick={addWorksheet} type="button"><Plus /></button>
            {sheetNames.map((sheet) => <button aria-label={sheet} className={`sheet-tab${activeSheet === sheet ? " active" : ""}`} key={sheet} onClick={() => switchSheet(sheet)} onMouseDown={(event) => { if (event.button !== 2) return; event.preventDefault(); event.stopPropagation(); deleteWorksheet(sheet, event.clientX, event.clientY); }} onContextMenu={(event) => { event.preventDefault(); event.stopPropagation(); deleteWorksheet(sheet, event.clientX, event.clientY); }} type="button">{sheet}</button>)}
            <span className="sheet-tab-spacer" />
          </div>
          <section className="table-card excel-table-card" onCopy={handleCopyEvent} onCut={handleCutEvent} onKeyDown={handleTableKeyDown} onPaste={handlePasteEvent} ref={tableCardRef} tabIndex={0}>
            <VirtualGrid cellSelected={isCellSelected} columnWidths={columnWidths} editRequest={editRequest} fields={activeFields} flashes={flashes} freeze={freeze} formulaContext={{ rows: state.rows, fields: FIELDS }} formulaReferences={formulaReferences} identity={identity} mergeGeometry={mergeGeometry} onAll={selectAll} onColumnResize={beginColumnResize} onColumnSelect={selectColumn} onColumnDrag={dragColumnSelection} onCommit={commitCell} onContext={(event, row, col) => { event.preventDefault(); if (!isCellSelected(row, col)) selectCell(row, col, false); setContextMenu({ row, col, x: event.clientX, y: event.clientY }); }} onDrag={dragSelectionTo} onFill={fillSelection} onFormulaDraftChange={setCellFormulaDraft} onFormulaEditorChange={registerFormulaEditor} onFormulaReferencePointerDown={formulaReferencePointerDown} onFormulaReferencePointerMove={formulaReferencePointerMove} onFocusCell={setFocusedCell} onMove={moveActiveCell} onRowDrag={dragRowSelection} onRowResize={beginRowResize} onRowSelect={selectRow} onScroll={handleGridScroll} onSelect={beginCellSelection} onSelectRange={selectMergedRange} rows={displayedRows} searchMatches={findMatchKeys} selectedBounds={selectedBounds} selection={selection} selections={allSelections} state={state} zoom={zoom} />
            <footer className="table-footer"><p /><p>{activeSheet}　 ·　 {selectedCellCount} 个单元格</p></footer>
          </section>
          <div className="sheet-statistics">
            <div className="sheet-statistics-values" aria-live="polite">
              <span>平均值={selectionStatistics.average}</span>
              <span>计数={selectionStatistics.count}</span>
              <span>求和={selectionStatistics.sum}</span>
            </div>
            <div className="sheet-zoom-controls" aria-label="缩放">
              <output className="sheet-zoom-value" aria-label="当前缩放比例">{zoom}%</output>
              <button aria-label="缩小" title="缩小" onClick={() => setZoom((current) => Math.max(50, current - 10))} type="button"><Minus /></button>
              <input aria-label="缩放比例" max={200} min={50} onChange={(event) => setZoom(Number(event.target.value))} step={10} style={{ "--zoom-progress": `${((zoom - 50) / 150) * 100}%` } as CSSProperties} type="range" value={zoom} />
              <button aria-label="放大" title="放大" onClick={() => setZoom((current) => Math.min(200, current + 10))} type="button"><Plus /></button>
            </div>
          </div>
        </section>
      </main>
      {menuOpen ? <div className="document-menu" role="menu"><button onClick={() => { exportCsv(); setMenuOpen(false); }} type="button">导出 CSV</button><button onClick={() => { void refresh(); setMenuOpen(false); }} type="button">刷新文档</button><button onClick={() => { void copyShareLink(); setMenuOpen(false); }} type="button">复制分享链接</button></div> : null}
      {identityOpen ? <IdentityDialog canClose={Boolean(identity.name)} currentName={identity.name} onClose={() => setIdentityOpen(false)} onSave={saveIdentity} /> : null}
      {contextMenu ? <SpreadsheetContextMenu menu={contextMenu} meta={state.rows[contextMenu.row] && FIELDS[contextMenu.col] ? state.cells[`${state.rows[contextMenu.row].id}:${FIELDS[contextMenu.col].key}`] : undefined} onClear={() => void clearSelection()} onClose={() => setContextMenu(null)} onCopy={() => void copySelection()} onCut={() => { setContextMenu(null); void cutSelection(); }} onDeleteRows={() => { setContextMenu(null); void deleteSelectedRows(); }} onInsertAbove={() => { setContextMenu(null); void insertSelectedRow("above"); }} onInsertBelow={() => { setContextMenu(null); void insertSelectedRow("below"); }} onMerge={() => { setContextMenu(null); void mergeSelection(); }} onPaste={() => void pasteFromSystemClipboard()} onResetColumnWidth={() => { setContextMenu(null); void resetSelectedDimensions("column"); }} onResetRowHeight={() => { setContextMenu(null); void resetSelectedDimensions("row"); }} selection={selection} onUnmerge={() => { setContextMenu(null); void unmergeSelection(); }} /> : null}
      {sheetContextMenu ? (
        <SheetContextMenu
          menu={sheetContextMenu}
          onClose={() => setSheetContextMenu(null)}
          onDelete={(sheetName) => { setSheetContextMenu(null); removeWorksheet(sheetName); }}
          onRename={(sheetName) => {
            const nextName = window.prompt("重命名工作表", sheetName);
            setSheetContextMenu(null);
            if (nextName !== null) renameWorksheet(sheetName, nextName);
          }}
        />
      ) : null}
      {toast ? <div className="toast" role="status">{toast}</div> : null}
    </div>
  );

  return (
    <div
      className="app-shell"
      onContextMenu={(event) => event.preventDefault()}
      onMouseDown={() => setContextMenu(null)}
    >
      <header className="topbar">
        <div className="brand-block">
          <div className="brand-mark" aria-hidden="true"><span /><span /><span /><span /></div>
          <div>
                  <p>GRIDBOOK</p>
            <h1>共享表格</h1>
          </div>
          <span className="demo-pill">DEMO</span>
        </div>
        <div className="topbar-right">
          <div className={`connection-state state-${connection}`}>
            <i />
            {connection === "live" ? "已实时同步" : connection === "connecting" ? "正在连接" : "正在重连"}
          </div>
          <div className="presence-stack" aria-label={`${onlineUsers.length} 人在线`}>
            {onlineUsers.slice(0, 4).map((person) => (
              <Avatar clientId={person.client_id} key={person.client_id} name={person.name} />
            ))}
            {onlineUsers.length > 4 ? <span className="avatar avatar-more">+{onlineUsers.length - 4}</span> : null}
          </div>
          <span className="online-copy">{onlineUsers.length || 1} 人在线</span>
          <button className="identity-button" onClick={() => setIdentityOpen(true)} type="button">
            <Avatar clientId={identity.clientId} name={identity.name || "访客"} />
            <span>{identity.name || "设置名字"}</span>
          </button>
        </div>
      </header>

      <main className="workspace">
        <section className="workspace-main">
          <div className="page-heading">
            <div>
              <div className="breadcrumb"><span>工作台</span><b>/</b><strong>共享表格</strong></div>
              <h2>共享表格</h2>
              <p>单击选择、双击编辑；支持从 Excel / WPS 直接复制粘贴多行多列。</p>
            </div>
            <div className="heading-actions">
              <button className="button secondary" onClick={exportCsv} type="button">
                <Icon name="download" />导出 CSV
              </button>
              <button className="button primary" disabled={addingRow || !identity.name} onClick={() => void handleAddRow()} type="button">
                <Icon name="plus" />{addingRow ? "正在新增" : "新增一行"}
              </button>
            </div>
          </div>

          <div className="scope-bar">
            <div className="scope-tabs">
              <button className="active" type="button">工作表</button>
              <button disabled type="button" title="本演示仅保留基础表格能力">视图</button>
            </div>
              <div className="period-chip">
                <span>共享表格</span>
                <strong>多人实时编辑</strong>
              </div>
            <div className="plain-data-note"><i />兼容 XLS / XLSX 剪贴板 · 无公式</div>
          </div>

          <div className="spreadsheet-toolbar" aria-label="表格编辑工具栏">
            <div className="toolbar-label">编辑</div>
            <div className="toolbar-group">
              <button onClick={() => void cutSelection()} title="剪切（Ctrl+X）" type="button">剪切</button>
              <button onClick={() => void copySelection()} title="复制（Ctrl+C）" type="button">复制</button>
              <button onClick={() => void pasteFromSystemClipboard()} title="粘贴（Ctrl+V）" type="button">粘贴</button>
              <button onClick={() => void clearSelection()} title="清空（Delete）" type="button">清空</button>
            </div>
            <span className="toolbar-separator" />
            <div className="toolbar-group formatting-group">
              <button
                className={activeFormat.bold ? "active" : ""}
                onClick={() => void applySelectedFormat({ bold: !activeFormat.bold })}
                title="粗体"
                type="button"
              ><b>B</b></button>
              <button
                className={activeFormat.italic ? "active" : ""}
                onClick={() => void applySelectedFormat({ italic: !activeFormat.italic })}
                title="斜体"
                type="button"
              ><i>I</i></button>
              <button
                className={activeFormat.wrap ? "active" : ""}
                onClick={() => void applySelectedFormat({ wrap: !activeFormat.wrap })}
                title="自动换行"
                type="button"
              >换行</button>
              {(["left", "center", "right"] as const).map((align) => (
                <button
                  className={activeFormat.align === align ? "active alignment-button" : "alignment-button"}
                  key={align}
                  onClick={() => void applySelectedFormat({ align })}
                  title={align === "left" ? "左对齐" : align === "center" ? "居中" : "右对齐"}
                  type="button"
                ><span className={`align-glyph align-${align}`}><i /><i /><i /></span></button>
              ))}
            </div>
            <span className="toolbar-separator" />
            <div className="toolbar-group fill-group" aria-label="填充颜色">
              {(["none", "yellow", "blue", "green", "rose"] as const).map((fill) => (
                <button
                  aria-label={`填充颜色 ${fill}`}
                  className={`fill-button fill-${fill}${activeFormat.fill === fill || (fill === "none" && !activeFormat.fill) ? " active" : ""}`}
                  key={fill}
                  onClick={() => void applySelectedFormat({ fill })}
                  title={fill === "none" ? "无填充" : "设置填充颜色"}
                  type="button"
                ><span /></button>
              ))}
            </div>
            <span className="toolbar-separator" />
            <div className="toolbar-group">
              <button
                className={selectedMerges.length ? "active" : ""}
                onClick={() => void (selectedMerges.length ? unmergeSelection() : mergeSelection())}
                title={selectedMerges.length ? "取消合并" : "合并选中单元格"}
                type="button"
              >{selectedMerges.length ? "取消合并" : "合并单元格"}</button>
              <button onClick={() => void insertSelectedRow("below")} title="在选区下方插入行" type="button">插入行</button>
              <button onClick={() => void deleteSelectedRows()} title="删除选中的整行" type="button">删除行</button>
            </div>
            <div className="toolbar-tip">拖动列标题右缘或行号下缘调整尺寸</div>
          </div>

          <section
            className="table-card"
            onCopy={handleCopyEvent}
            onCut={handleCutEvent}
            onKeyDown={handleTableKeyDown}
            onPaste={handlePasteEvent}
            ref={tableCardRef}
            tabIndex={0}
          >
            <div className="table-titlebar">
              <div>
                <h3>表格内容</h3>
              <span>{state.rows.length} 行</span>
              </div>
              <div className="table-hint"><Icon name="users" />Shift 扩选 · 右键打开表格菜单</div>
            </div>
            <VirtualGrid
              cellSelected={isCellSelected}
              columnWidths={columnWidths}
              editRequest={editRequest}
              fields={FIELDS}
              flashes={flashes}
              freeze={freeze}
              formulaContext={{ rows: state.rows, fields: FIELDS }}
              identity={identity}
              mergeGeometry={mergeGeometry}
              onAll={selectAll}
              onColumnResize={beginColumnResize}
              onColumnSelect={selectColumn}
              onColumnDrag={dragColumnSelection}
              onCommit={commitCell}
              onContext={(event, row, col) => {
                event.preventDefault();
                event.stopPropagation();
                if (!isCellSelected(row, col)) selectCell(row, col, false);
                setContextMenu({ row, col, x: event.clientX, y: event.clientY });
              }}
              onDrag={dragSelectionTo}
              onFill={fillSelection}
              formulaReferences={formulaReferences}
              onFormulaEditorChange={registerFormulaEditor}
              onFormulaDraftChange={setCellFormulaDraft}
              onFormulaReferencePointerDown={formulaReferencePointerDown}
              onFormulaReferencePointerMove={formulaReferencePointerMove}
              onFocusCell={setFocusedCell}
              onMove={moveActiveCell}
              onRowDrag={dragRowSelection}
              onRowResize={beginRowResize}
              onRowSelect={selectRow}
              onScroll={handleGridScroll}
              onSelect={beginCellSelection}
              onSelectRange={selectMergedRange}
              rows={state.rows}
              searchMatches={findMatchKeys}
              selectedBounds={selectedBounds}
              selection={selection}
              selections={allSelections}
              state={state}
              zoom={zoom}
            />
            {false && <div className="table-scroll">
              <table className="contribution-table">
                <colgroup>
                  <col className="row-number-column" style={{ width: 46 }} />
                  {FIELDS.map((field, index) => <col key={field.key} style={{ width: columnWidths[index] }} />)}
                </colgroup>
                <thead>
                  <tr className="column-letter-row">
                    <th
                      aria-label="选择全部"
                      className="sheet-corner"
                      onMouseDown={(event) => {
                        event.preventDefault();
                        selectAll();
                      }}
                    ><span /></th>
                    {FIELDS.map((field, columnIndex) => (
                      <th
                        className={columnIndex >= selectedBounds.left && columnIndex <= selectedBounds.right ? "header-selected" : ""}
                        key={field.key}
                        onMouseDown={(event) => {
                          event.preventDefault();
                          selectColumn(columnIndex, event.shiftKey);
                        }}
                      >
                        {excelColumnName(columnIndex)}
                      </th>
                    ))}
                  </tr>
                  <tr className="field-label-row">
                    <th className="row-heading">#</th>
                    {FIELDS.map((field, columnIndex) => (
                      <th
                        className={columnIndex >= selectedBounds.left && columnIndex <= selectedBounds.right ? "header-selected" : ""}
                        key={field.key}
                        onMouseDown={(event) => {
                          event.preventDefault();
                          selectColumn(columnIndex, event.shiftKey);
                        }}
                      >
                        <span>{field.shortLabel ?? field.label}</span>
                        <i
                          aria-hidden="true"
                          className="column-resizer"
                          onDoubleClick={(event) => {
                            event.preventDefault();
                            event.stopPropagation();
                            setColumnWidths((current) => current.map((width, index) => index === columnIndex ? DEFAULT_WORKSHEET_COLUMN_WIDTH : width));
                          }}
                          onMouseDown={(event) => beginColumnResize(columnIndex, event)}
                        />
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {loading ? (
                    <tr><td className="table-message" colSpan={FIELDS.length + 1}><span className="loading-ring" />正在载入协作表…</td></tr>
                  ) : state.rows.length === 0 ? (
                    <tr><td className="table-message" colSpan={FIELDS.length + 1}>表格为空，点击“新增一行”开始填写。</td></tr>
                  ) : (
                    state.rows.map((row, rowIndex) => (
                      <tr
                        key={row.id}
                        style={{ "--row-height": `${state.layout.row_heights[String(row.id)] ?? 55}px` } as CSSProperties}
                      >
                        <th
                          className={`row-number${rowIndex >= selectedBounds.top && rowIndex <= selectedBounds.bottom ? " header-selected" : ""}`}
                          onMouseDown={(event) => {
                            event.preventDefault();
                            selectRow(rowIndex, event.shiftKey);
                          }}
                          scope="row"
                        >
                          {rowIndex + 1}
                          <i
                            aria-hidden="true"
                            className="row-resizer"
                            onDoubleClick={(event) => {
                              event.preventDefault();
                              event.stopPropagation();
                              void updateLayout(identity, "row_height", { row_id: row.id, height: 55 });
                            }}
                            onMouseDown={(event) => beginRowResize(row, event)}
                          />
                        </th>
                        {FIELDS.map((field, columnIndex) => {
                          const coordinateKey = `${rowIndex}:${columnIndex}`;
                          if (mergeGeometry.hidden.has(coordinateKey)) return null;
                          const mergeInfo = mergeGeometry.anchors.get(coordinateKey);
                          const key = `${row.id}:${field.key}`;
                          const cellSelected = isCellSelected(rowIndex, columnIndex);
                          const cellFormat = state.layout.formats[key] ?? {};
                          const mergeBottom = rowIndex + (mergeInfo?.rowSpan ?? 1) - 1;
                          const mergeRight = columnIndex + (mergeInfo?.colSpan ?? 1) - 1;
                          const remoteEditor = state.presence.find(
                            (person) =>
                              person.client_id !== identity.clientId &&
                              person.cell?.row_id === row.id &&
                              person.cell.field === field.key,
                          );
                          return (
                            <td
                              className={[
                                `field-${field.key}`,
                                mergeInfo ? "merged-cell" : "",
                                cellFormat.bold ? "format-bold" : "",
                                cellFormat.italic ? "format-italic" : "",
                                cellFormat.wrap ? "format-wrap" : "",
                                cellFormat.align ? `format-align-${cellFormat.align}` : "",
                                cellFormat.fill && cellFormat.fill !== "none" ? `format-fill-${cellFormat.fill}` : "",
                              ].filter(Boolean).join(" ")}
                              colSpan={mergeInfo?.colSpan}
                              key={field.key}
                              rowSpan={mergeInfo?.rowSpan}
                            >
                              <EditableCell
                                field={field}
                                flash={flashes.has(key)}
                                identity={identity}
                                meta={state.cells[key]}
                                active={selection.focus.row === rowIndex && selection.focus.col === columnIndex}
                                editRequest={
                                  editRequest?.row === rowIndex && editRequest.col === columnIndex
                                    ? { id: editRequest.id, seed: editRequest.seed }
                                    : undefined
                                }
                                onCommit={(value, version) => commitCell(row.id, field.key, value, version)}
                                onFocusCell={setFocusedCell}
                                onMoveAfterCommit={moveActiveCell}
                                onOpenContextMenu={(event) => {
                                  event.preventDefault();
                                  event.stopPropagation();
                                  if (!cellSelected) selectCell(rowIndex, columnIndex, false);
                                  setContextMenu({
                                    row: rowIndex,
                                    col: columnIndex,
                                    x: event.clientX,
                                    y: event.clientY,
                                  });
                                }}
                                onSelect={(extend) => mergeInfo
                                  ? selectMergedRange(rowIndex, columnIndex, mergeInfo.rowSpan, mergeInfo.colSpan, extend)
                                  : beginCellSelection(rowIndex, columnIndex, extend)}
                                onSelectionDrag={() => dragSelectionTo(rowIndex, columnIndex)}
                                remoteEditor={remoteEditor}
                                row={row}
                                selected={cellSelected}
                                selectionEdges={{
                                  top: cellSelected && rowIndex === selectedBounds.top,
                                  bottom: cellSelected && mergeBottom === selectedBounds.bottom,
                                  left: cellSelected && columnIndex === selectedBounds.left,
                                  right: cellSelected && mergeRight === selectedBounds.right,
                                }}
                              />
                            </td>
                          );
                        })}
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>}
            <footer className="table-footer">
              <p><span className="autosave-dot" />{pasting ? "正在批量粘贴…" : "所有更改自动保存"}<b>已选择 {selectedCellCount} 格</b></p>
              <p>{latestUpdate ? `最近更新 ${fullTime(latestUpdate)}` : "等待首次编辑"}<span>·</span>版本 {state.revision}</p>
            </footer>
          </section>

          <div className="environment-note">
            <span className="environment-icon"><Icon name="check" /></span>
            <div>
              <strong>这是独立的多人编辑 Demo</strong>
              <p>这是独立的共享表格，编辑记录和导出数据均与正式项目隔离。</p>
            </div>
          </div>
        </section>

        <ActivityFeed items={state.activity} />
      </main>

      {identityOpen ? (
        <IdentityDialog
          canClose={Boolean(identity.name)}
          currentName={identity.name}
          onClose={() => setIdentityOpen(false)}
          onSave={saveIdentity}
        />
      ) : null}
      {contextMenu ? (
        <SpreadsheetContextMenu
          menu={contextMenu!}
          meta={
            state.rows[contextMenu!.row] && FIELDS[contextMenu!.col]
              ? state.cells[`${state.rows[contextMenu!.row].id}:${FIELDS[contextMenu!.col].key}`]
              : undefined
          }
          onClear={() => void clearSelection()}
          onClose={() => setContextMenu(null)}
          onCopy={() => void copySelection()}
          onCut={() => { setContextMenu(null); void cutSelection(); }}
          onDeleteRows={() => { setContextMenu(null); void deleteSelectedRows(); }}
          onInsertAbove={() => { setContextMenu(null); void insertSelectedRow("above"); }}
          onInsertBelow={() => { setContextMenu(null); void insertSelectedRow("below"); }}
          onMerge={() => { setContextMenu(null); void mergeSelection(); }}
          onPaste={() => void pasteFromSystemClipboard()}
          onResetColumnWidth={() => { setContextMenu(null); void resetSelectedDimensions("column"); }}
          onResetRowHeight={() => { setContextMenu(null); void resetSelectedDimensions("row"); }}
          selection={selection}
          onUnmerge={() => { setContextMenu(null); void unmergeSelection(); }}
        />
      ) : null}
      {sheetContextMenu ? (
        <div
          className="sheet-context-menu"
          onMouseDown={(event) => event.stopPropagation()}
          role="menu"
          style={{ left: Math.min(sheetContextMenu!.x, window.innerWidth - 190), top: Math.min(sheetContextMenu!.y, window.innerHeight - 110) }}
        >
          <div className="sheet-context-menu-title">{sheetContextMenu!.sheet}</div>
          <button onClick={() => { const nextName = window.prompt("重命名工作表", sheetContextMenu!.sheet); setSheetContextMenu(null); if (nextName !== null) renameWorksheet(sheetContextMenu!.sheet, nextName); }} role="menuitem" type="button">重命名</button>
          <button onClick={() => { setSheetContextMenu(null); removeWorksheet(sheetContextMenu!.sheet); }} role="menuitem" type="button">删除工作表</button>
        </div>
      ) : null}
      {toast ? <div className="toast" role="status"><Icon name="check" />{toast}</div> : null}
    </div>
  );
}

export default InfiniteWorkbook;
