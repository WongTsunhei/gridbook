import type { EditableField, GridBookRow } from "./types";

export interface FormulaField {
  key: EditableField;
}

export interface FormulaContext {
  rows: GridBookRow[];
  fields: FormulaField[];
}

type FormulaValue = number | string | boolean | null | FormulaValue[];

const CELL_REFERENCE = /^(?:(?:'[^']+'|[A-Za-z_\u4e00-\u9fff][\w\u4e00-\u9fff]*)!)?\$?[A-Za-z]{1,3}\$?\d+/;

function flatten(value: FormulaValue): FormulaValue[] {
  return Array.isArray(value) ? value.flatMap(flatten) : [value];
}

function numeric(value: FormulaValue): number | null {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value === "boolean") return value ? 1 : 0;
  if (typeof value === "string" && value.trim() !== "") {
    const parsed = Number(value.replaceAll(",", "").replace(/%$/, ""));
    return Number.isFinite(parsed) ? parsed : null;
  }
  return null;
}

function truthy(value: FormulaValue): boolean {
  if (typeof value === "boolean") return value;
  const parsed = numeric(value);
  return parsed === null ? Boolean(value) : parsed !== 0;
}

function columnIndex(label: string): number {
  let result = 0;
  for (const char of label.toUpperCase()) result = result * 26 + char.charCodeAt(0) - 64;
  return result - 1;
}

function columnLabel(index: number): string {
  let value = index + 1;
  let label = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

function parseReference(reference: string): { sheet: string | null; col: number; row: number } | null {
  const match = reference.match(/^(?:(?:'([^']+)'|([A-Za-z_\u4e00-\u9fff][\w\u4e00-\u9fff]*))!)?\$?([A-Za-z]{1,3})\$?(\d+)$/);
  if (!match) return null;
  return { sheet: match[1] ?? match[2] ?? null, col: columnIndex(match[3]), row: Number(match[4]) - 1 };
}

function readCell(reference: string, context: FormulaContext, stack: Set<string>): FormulaValue {
  const parsed = parseReference(reference);
  if (!parsed || parsed.row < 0 || parsed.col < 0 || parsed.row >= context.rows.length || parsed.col >= context.fields.length) return "#REF!";
  const key = `${parsed.row}:${parsed.col}`;
  if (stack.has(key)) return "#CIRCULAR!";
  const raw = context.rows[parsed.row][context.fields[parsed.col].key];
  if (typeof raw === "string" && raw.trim().startsWith("=")) {
    const next = new Set(stack);
    next.add(key);
    return evaluateFormula(raw, context, next);
  }
  return raw === undefined ? null : (raw as FormulaValue);
}

function readRange(start: string, end: string, context: FormulaContext, stack: Set<string>): FormulaValue[] {
  const first = parseReference(start);
  const last = parseReference(end);
  if (!first || !last) return [];
  const values: FormulaValue[] = [];
  for (let row = Math.min(first.row, last.row); row <= Math.max(first.row, last.row); row += 1) {
    for (let col = Math.min(first.col, last.col); col <= Math.max(first.col, last.col); col += 1) {
      values.push(readCell(`${columnLabel(col)}${row + 1}`, context, stack));
    }
  }
  return values;
}

class FormulaParser {
  private index = 0;
  constructor(private readonly source: string, private readonly context: FormulaContext, private readonly stack: Set<string>) {}

  parse(): FormulaValue {
    const value = this.parseComparison();
    this.skipWhitespace();
    return this.index === this.source.length ? value : "#VALUE!";
  }

  private skipWhitespace() { while (/\s/.test(this.source[this.index] ?? "")) this.index += 1; }
  private match(text: string): boolean {
    this.skipWhitespace();
    if (this.source.slice(this.index, this.index + text.length) !== text) return false;
    this.index += text.length;
    return true;
  }

  private parseComparison(): FormulaValue {
    let left = this.parseAdditive();
    for (;;) {
      this.skipWhitespace();
      const operator = ["<>", "<=", ">=", "=", "<", ">"].find((item) => this.source.startsWith(item, this.index));
      if (!operator) return left;
      this.index += operator.length;
      const right = this.parseAdditive();
      const leftNumber = numeric(left);
      const rightNumber = numeric(right);
      const a = leftNumber !== null && rightNumber !== null ? leftNumber : String(left ?? "");
      const b = leftNumber !== null && rightNumber !== null ? rightNumber : String(right ?? "");
      left = operator === "=" ? a === b : operator === "<>" ? a !== b : operator === "<" ? a < b : operator === ">" ? a > b : operator === "<=" ? a <= b : a >= b;
    }
  }

  private parseAdditive(): FormulaValue {
    let left = this.parseMultiplicative();
    for (;;) {
      if (this.match("+")) left = (numeric(left) ?? 0) + (numeric(this.parseMultiplicative()) ?? 0);
      else if (this.match("-")) left = (numeric(left) ?? 0) - (numeric(this.parseMultiplicative()) ?? 0);
      else return left;
    }
  }

  private parseMultiplicative(): FormulaValue {
    let left = this.parseUnary();
    for (;;) {
      if (this.match("*")) left = (numeric(left) ?? 0) * (numeric(this.parseUnary()) ?? 0);
      else if (this.match("/")) {
        const divisor = numeric(this.parseUnary());
        if (!divisor) return "#DIV/0!";
        left = (numeric(left) ?? 0) / divisor;
      } else return left;
    }
  }

  private parseUnary(): FormulaValue {
    if (this.match("+")) return numeric(this.parseUnary()) ?? 0;
    if (this.match("-")) return -(numeric(this.parseUnary()) ?? 0);
    return this.parsePrimary();
  }

  private parsePrimary(): FormulaValue {
    this.skipWhitespace();
    if (this.match("(")) {
      const value = this.parseComparison();
      if (!this.match(")")) return "#VALUE!";
      return value;
    }
    const stringMatch = this.source.slice(this.index).match(/^"((?:[^"]|"")*)"/);
    if (stringMatch) { this.index += stringMatch[0].length; return stringMatch[1].replaceAll('""', '"'); }
    const numberMatch = this.source.slice(this.index).match(/^(?:\d+(?:\.\d*)?|\.\d+)/);
    if (numberMatch) { this.index += numberMatch[0].length; return Number(numberMatch[0]); }
    const refMatch = this.source.slice(this.index).match(CELL_REFERENCE);
    if (refMatch) {
      this.index += refMatch[0].length;
      if (this.match(":")) {
        const end = this.source.slice(this.index).match(CELL_REFERENCE);
        if (!end) return "#VALUE!";
        this.index += end[0].length;
        return readRange(refMatch[0], end[0], this.context, this.stack);
      }
      return readCell(refMatch[0], this.context, this.stack);
    }
    const identifier = this.source.slice(this.index).match(/^[A-Za-z_\u4e00-\u9fff][\w\u4e00-\u9fff]*/);
    if (!identifier) return "#VALUE!";
    this.index += identifier[0].length;
    if (!this.match("(")) return identifier[0].toUpperCase() === "TRUE";
    const args: FormulaValue[] = [];
    this.skipWhitespace();
    if (!this.match(")")) {
      for (;;) {
        args.push(this.parseComparison());
        if (this.match(")")) break;
        if (!this.match(",") && !this.match(";")) return "#VALUE!";
      }
    }
    return this.call(identifier[0].toUpperCase(), args);
  }

  private call(name: string, args: FormulaValue[]): FormulaValue {
    const values = args.flatMap(flatten);
    const numbers = values.map(numeric).filter((value): value is number => value !== null);
    switch (name) {
      case "SUM": return numbers.reduce((sum, value) => sum + value, 0);
      case "AVERAGE": return numbers.length ? numbers.reduce((sum, value) => sum + value, 0) / numbers.length : 0;
      case "COUNT": return numbers.length;
      case "COUNTA": return values.filter((value) => value !== null && value !== "").length;
      case "MAX": return numbers.length ? Math.max(...numbers) : 0;
      case "MIN": return numbers.length ? Math.min(...numbers) : 0;
      case "IF": return truthy(args[0] ?? false) ? (args[1] ?? false) : (args[2] ?? false);
      case "AND": return args.every(truthy);
      case "OR": return args.some(truthy);
      case "NOT": return !truthy(args[0] ?? false);
      case "ROUND": { const factor = 10 ** (numeric(args[1] ?? 0) ?? 0); return Math.round((numeric(args[0] ?? 0) ?? 0) * factor) / factor; }
      case "ROUNDUP": { const factor = 10 ** (numeric(args[1] ?? 0) ?? 0); const value = numeric(args[0] ?? 0) ?? 0; return Math.ceil(value * factor) / factor; }
      case "ROUNDDOWN": { const factor = 10 ** (numeric(args[1] ?? 0) ?? 0); const value = numeric(args[0] ?? 0) ?? 0; return Math.floor(value * factor) / factor; }
      default: return "#NAME?";
    }
  }
}

export function evaluateFormula(value: string, context: FormulaContext, stack = new Set<string>()): FormulaValue {
  const source = value.trim().startsWith("=") ? value.trim().slice(1) : value.trim();
  if (!source) return "";
  try { return new FormulaParser(source, context, stack).parse(); } catch { return "#VALUE!"; }
}

export function displayFormulaValue(value: unknown, context: FormulaContext): unknown {
  if (typeof value !== "string" || !value.trim().startsWith("=")) return value;
  return evaluateFormula(value, context);
}
