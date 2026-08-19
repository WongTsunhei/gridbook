import type { CellFormat, CollaborationIdentity } from "./types";

const API_BASE = `${import.meta.env.BASE_URL}api/v2`;

export interface V2CellMeta {
  version: number;
  updated_by: string;
  updated_at: string;
  client_id: string;
}

export interface V2Merge {
  id: string;
  start: { row: number; col: number };
  end: { row: number; col: number };
}

export interface V2UsedRange {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

export interface V2SheetSummary { id: string; name: string; }

export interface V2SheetWindow {
  id: string;
  name: string;
  used_range: V2UsedRange;
  content_range: V2UsedRange;
  cells: Record<string, string>;
  meta: Record<string, V2CellMeta>;
  formats: Record<string, CellFormat>;
  row_heights: Record<string, number>;
  column_widths: Record<string, number>;
  merges: V2Merge[];
}

export interface V2Snapshot {
  schema_version: 2;
  revision: number;
  sheets: V2SheetSummary[];
  sheet: V2SheetWindow;
  window: V2UsedRange;
  presence?: Array<{ client_id: string; name: string; sheet_id: string; cell: { row: number; col: number } | null }>;
}

export interface V2Delta {
  kind: "delta";
  sheet_id: string;
  revision: number;
  cells: Array<{ row: number; col: number; value: string; meta: V2CellMeta }>;
  formats: Array<{ row: number; col: number; format: CellFormat }>;
  row_heights: Record<string, number>;
  column_widths: Record<string, number>;
  merges: V2Merge[] | null;
  structure_changed: boolean;
  used_range: V2UsedRange;
  content_range: V2UsedRange;
}

export class WorkbookV2ApiError extends Error {
  constructor(public readonly status: number, message: string) { super(message); }
}

async function json<T>(path: string, init?: RequestInit): Promise<T> {
  const response = await fetch(`${API_BASE}${path}`, {
    ...init,
    headers: init?.body ? { "Content-Type": "application/json", ...init.headers } : init?.headers,
  });
  const text = await response.text();
  let payload: Record<string, unknown>;
  try { payload = JSON.parse(text) as Record<string, unknown>; }
  catch { throw new WorkbookV2ApiError(response.status || 502, text.slice(0, 180) || "接口返回格式错误"); }
  if (!response.ok) throw new WorkbookV2ApiError(response.status, String(payload.error ?? "请求失败"));
  return payload as T;
}

function sheetPath(id: string): string { return `/sheets/${encodeURIComponent(id)}`; }

export function getV2Sheets(): Promise<{ revision: number; sheets: V2SheetSummary[] }> {
  return json("/sheets");
}

export function getV2Window(sheetId: string, top: number, left: number, rows: number, cols: number): Promise<V2Snapshot> {
  const query = new URLSearchParams({ top: String(top), left: String(left), rows: String(rows), cols: String(cols) });
  return json(`${sheetPath(sheetId)}/window?${query}`);
}

export function v2EventsUrl(identity: CollaborationIdentity, sheetId: string, top = 0, left = 0, rows = 64, cols = 24): string {
  const query = new URLSearchParams({
    client_id: identity.clientId,
    name: identity.name,
    sheet_id: sheetId,
    top: String(top),
    left: String(left),
    rows: String(rows),
    cols: String(cols),
  });
  return `${API_BASE}/events?${query}`;
}

export function createV2Sheet(identity: CollaborationIdentity, name: string): Promise<{ sheet: V2SheetSummary; revision: number; sheets: V2SheetSummary[] }> {
  return json("/sheets", { method: "POST", body: JSON.stringify({ client_id: identity.clientId, name: identity.name, action: "create", sheet_name: name }) });
}

export function renameV2Sheet(identity: CollaborationIdentity, sheetId: string, name: string): Promise<{ revision: number; sheets: V2SheetSummary[] }> {
  return json(`${sheetPath(sheetId)}/rename`, { method: "POST", body: JSON.stringify({ client_id: identity.clientId, name: identity.name, sheet_name: name }) });
}

export function deleteV2Sheet(identity: CollaborationIdentity, sheetId: string): Promise<{ revision: number; sheets: V2SheetSummary[] }> {
  return json(sheetPath(sheetId), { method: "DELETE", body: JSON.stringify({ client_id: identity.clientId, name: identity.name }) });
}

export function applyV2Operations(identity: CollaborationIdentity, sheetId: string, operations: Array<Record<string, unknown>>, cell?: { row: number; col: number }): Promise<V2Delta> {
  return json(`${sheetPath(sheetId)}/ops`, { method: "POST", body: JSON.stringify({ client_id: identity.clientId, name: identity.name, operations, cell }) });
}

export function updateV2Presence(identity: CollaborationIdentity, sheetId: string, cell: { row: number; col: number } | null): Promise<{ presence: V2Snapshot["presence"] }> {
  return json("/presence", { method: "POST", body: JSON.stringify({ client_id: identity.clientId, name: identity.name, sheet_id: sheetId, cell }) });
}
