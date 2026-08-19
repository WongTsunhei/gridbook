import type { CellFormat } from "./types";

export type FillValue = string | number | null;

export interface FillBounds {
  top: number;
  bottom: number;
  left: number;
  right: number;
}

export interface FillGridRow {
  id: number;
  [key: string]: unknown;
}

export interface FillGridField {
  key: string;
}

export interface FillGrid {
  rows: FillGridRow[];
  fields: FillGridField[];
  formats: Record<string, CellFormat>;
}

export interface FillCellResult {
  value: FillValue;
  format: CellFormat;
  sourceRow: number;
  sourceCol: number;
}

export type FillMode = "auto" | "copy";

const NUMBER_TOKEN = /^(.*?)(-?\d+(?:\.\d+)?)([^\d]*)$/;
const CELL_REFERENCE = /(\$?)([A-Z]{1,3})(\$?)(\d+)/g;
const DAY_NAMES = new Map([
  ["星期一", 0], ["星期二", 1], ["星期三", 2], ["星期四", 3], ["星期五", 4], ["星期六", 5], ["星期日", 6], ["星期天", 6],
  ["Monday", 0], ["Tuesday", 1], ["Wednesday", 2], ["Thursday", 3], ["Friday", 4], ["Saturday", 5], ["Sunday", 6],
]);
const MONTH_NAMES = new Map([
  ["一月", 0], ["二月", 1], ["三月", 2], ["四月", 3], ["五月", 4], ["六月", 5], ["七月", 6], ["八月", 7], ["九月", 8], ["十月", 9], ["十一月", 10], ["十二月", 11],
  ["January", 0], ["February", 1], ["March", 2], ["April", 3], ["May", 4], ["June", 5], ["July", 6], ["August", 7], ["September", 8], ["October", 9], ["November", 10], ["December", 11],
]);

function mod(value: number, divisor: number): number {
  return ((value % divisor) + divisor) % divisor;
}

function columnIndex(value: string): number {
  let output = 0;
  for (const char of value) output = output * 26 + char.charCodeAt(0) - 64;
  return output - 1;
}

function columnName(index: number): string {
  let value = index + 1;
  let output = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    output = String.fromCharCode(65 + remainder) + output;
    value = Math.floor((value - 1) / 26);
  }
  return output;
}

function asText(value: FillValue): string {
  return value === null || value === undefined ? "" : String(value);
}

function numericValue(value: FillValue): number | null {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value !== "string" || !value.trim()) return null;
  const parsed = Number(value.trim());
  return Number.isFinite(parsed) ? parsed : null;
}

function numberToken(value: FillValue): { prefix: string; suffix: string; number: number; width: number } | null {
  const text = asText(value);
  const match = text.match(NUMBER_TOKEN);
  if (!match) return null;
  const numeric = Number(match[2]);
  if (!Number.isFinite(numeric)) return null;
  const digits = match[2].replace("-", "").split(".")[0] ?? "";
  return { prefix: match[1], suffix: match[3], number: numeric, width: digits.length };
}

function dateValue(value: FillValue): { date: Date; format: "ymd" | "mdy" | "dmy" | "iso" } | null {
  if (typeof value !== "string") return null;
  const text = value.trim();
  const match = text.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
  if (!match) return null;
  const year = Number(match[1]);
  const month = Number(match[2]) - 1;
  const day = Number(match[3]);
  const date = new Date(Date.UTC(year, month, day));
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month || date.getUTCDate() !== day) return null;
  return { date, format: text.includes("-") ? "iso" : "ymd" };
}

function extendNumber(value: FillValue, step: number, offset: number): FillValue {
  const token = numberToken(value);
  if (!token) return value;
  const next = token.number + step * offset;
  const numeric = Number.isInteger(next) ? String(next) : String(Number(next.toFixed(10)));
  const [integer, fraction] = numeric.split(".");
  const padded = token.width > integer.replace("-", "").length
    ? `${integer.startsWith("-") ? "-" : ""}${integer.replace("-", "").padStart(token.width, "0")}`
    : integer;
  return `${token.prefix}${fraction ? `${padded}.${fraction}` : padded}${token.suffix}`;
}

function extendDate(value: FillValue, stepDays: number, offset: number): FillValue {
  const parsed = dateValue(value);
  if (!parsed) return value;
  const date = new Date(parsed.date.getTime());
  date.setUTCDate(date.getUTCDate() + stepDays * offset);
  const year = date.getUTCFullYear();
  const month = String(date.getUTCMonth() + 1).padStart(2, "0");
  const day = String(date.getUTCDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function inferStep(values: FillValue[]): { kind: "number" | "date" | "number-token" | "day" | "month"; step: number } | null {
  if (values.length === 0) return null;
  const dates = values.map(dateValue);
  if (dates.every(Boolean)) {
    if (values.length === 1) return { kind: "date", step: 1 };
    const first = dates[0]!.date.getTime();
    const second = dates[1]!.date.getTime();
    return { kind: "date", step: Math.round((second - first) / 86_400_000) || 1 };
  }
  const tokens = values.map(numberToken);
  if (tokens.every(Boolean) && new Set(tokens.map((item) => `${item!.prefix}\u0000${item!.suffix}`)).size === 1) {
    if (values.length === 1) return { kind: "number-token", step: 1 };
    return { kind: "number-token", step: tokens[1]!.number - tokens[0]!.number };
  }
  const numbers = values.map(numericValue);
  if (numbers.every((value) => value !== null)) {
    if (values.length === 1) return { kind: "number", step: 1 };
    return { kind: "number", step: numbers[1]! - numbers[0]! };
  }
  const dayIndexes = values.map((value) => DAY_NAMES.get(asText(value)) ?? -1);
  if (dayIndexes.every((value) => value >= 0)) {
    return { kind: "day", step: mod(dayIndexes[1] - dayIndexes[0], 7) || 1 };
  }
  const monthIndexes = values.map((value) => MONTH_NAMES.get(asText(value)) ?? -1);
  if (monthIndexes.every((value) => value >= 0)) {
    return { kind: "month", step: mod(monthIndexes[1] - monthIndexes[0], 12) || 1 };
  }
  return null;
}

function extendPattern(values: FillValue[], offset: number): FillValue | null {
  if (!values.length) return null;
  const inferred = inferStep(values);
  if (!inferred) return null;
  const base = values[0];
  const step = inferred.step;
  if (inferred.kind === "date") return extendDate(base, step, offset);
  if (inferred.kind === "number-token") return extendNumber(base, step, offset);
  if (inferred.kind === "number") {
    const numeric = numericValue(base);
    if (numeric === null) return null;
    return numeric + step * offset;
  }
  const text = asText(base);
  if (inferred.kind === "day" && DAY_NAMES.has(text)) {
    const index = DAY_NAMES.get(text)!;
    const next = (index + step * offset) % 7;
    const english = /^[A-Z]/.test(text);
    return [...DAY_NAMES.entries()].find(([label, value]) => value === mod(next, 7) && english === /^[A-Z]/.test(label))?.[0] ?? text;
  }
  if (inferred.kind === "month" && MONTH_NAMES.has(text)) {
    const index = MONTH_NAMES.get(text)!;
    const next = (index + step * offset) % 12;
    const english = /^[A-Z]/.test(text);
    return [...MONTH_NAMES.entries()].find(([label, value]) => value === mod(next, 12) && english === /^[A-Z]/.test(label))?.[0] ?? text;
  }
  return null;
}

export function translateFormula(formula: string, rowDelta: number, colDelta: number): string {
  return formula.replace(CELL_REFERENCE, (full, colAbsolute: string, colText: string, rowAbsolute: string, rowText: string) => {
    const nextColumn = colAbsolute ? columnIndex(colText) : columnIndex(colText) + colDelta;
    const nextRow = rowAbsolute ? Number(rowText) : Number(rowText) + rowDelta;
    if (nextColumn < 0 || nextRow < 1) return full;
    return `${colAbsolute ? "$" : ""}${columnName(nextColumn)}${rowAbsolute ? "$" : ""}${nextRow}`;
  });
}

function sourceCoordinate(bounds: FillBounds, row: number, col: number): { row: number; col: number } {
  return {
    row: bounds.top + mod(row - bounds.top, bounds.bottom - bounds.top + 1),
    col: bounds.left + mod(col - bounds.left, bounds.right - bounds.left + 1),
  };
}

function cellAt(grid: FillGrid, row: number, col: number): { value: FillValue; format: CellFormat } {
  const sourceRow = grid.rows[row];
  const field = grid.fields[col];
  if (!sourceRow || !field) return { value: null, format: {} };
  const raw = sourceRow[field.key];
  return {
    value: raw === undefined ? null : raw as FillValue,
    format: { ...(grid.formats[`${sourceRow.id}:${field.key}`] ?? {}) },
  };
}

function axisPattern(grid: FillGrid, bounds: FillBounds, row: number, col: number, axis: "row" | "column"): FillValue[] {
  if (axis === "column") return Array.from({ length: bounds.bottom - bounds.top + 1 }, (_, offset) => cellAt(grid, bounds.top + offset, col).value);
  return Array.from({ length: bounds.right - bounds.left + 1 }, (_, offset) => cellAt(grid, row, bounds.left + offset).value);
}

export function fillCellAt(grid: FillGrid, source: FillBounds, row: number, col: number, mode: FillMode = "auto"): FillCellResult {
  const mapped = sourceCoordinate(source, row, col);
  const base = cellAt(grid, mapped.row, mapped.col);
  if (mode === "copy") return { ...base, sourceRow: mapped.row, sourceCol: mapped.col };

  const rowOutside = row < source.top || row > source.bottom;
  const colOutside = col < source.left || col > source.right;
  const vertical = rowOutside && source.bottom >= source.top;
  const horizontal = colOutside && source.right >= source.left;
  const axis = vertical && source.bottom > source.top ? "column" : horizontal ? "row" : vertical ? "column" : "row";
  const axisIndex = axis === "column" ? mapped.col : mapped.row;
  const values = axisPattern(grid, source, axis === "column" ? source.top : mapped.row, axisIndex, axis);
  const offset = axis === "column"
    ? row > source.bottom ? row - source.bottom : row < source.top ? row - source.top : 0
    : col > source.right ? col - source.right : col < source.left ? col - source.left : 0;
  const formula = typeof base.value === "string" && base.value.trimStart().startsWith("=")
    ? translateFormula(base.value, row - mapped.row, col - mapped.col)
    : null;
  const generated = formula ?? (offset !== 0 ? extendPattern(values, offset > 0 ? offset + (values.length - 1) : offset) : null);
  return {
    value: generated === null ? base.value : generated,
    format: base.format,
    sourceRow: mapped.row,
    sourceCol: mapped.col,
  };
}

export function buildFillPreview(
  grid: FillGrid,
  source: FillBounds,
  target: FillBounds,
  mode: FillMode = "auto",
): Map<string, FillCellResult> {
  const output = new Map<string, FillCellResult>();
  for (let row = target.top; row <= target.bottom; row += 1) {
    const gridRow = grid.rows[row];
    if (!gridRow) continue;
    for (let col = target.left; col <= target.right; col += 1) {
      if (row >= source.top && row <= source.bottom && col >= source.left && col <= source.right) continue;
      const field = grid.fields[col];
      if (!field) continue;
      output.set(`${row}:${col}`, fillCellAt(grid, source, row, col, mode));
    }
  }
  return output;
}
