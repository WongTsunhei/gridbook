import type {
  CollaborationIdentity,
  CellBorders,
  CellFormat,
  EditableField,
  PresenceItem,
  GridBookState,
} from "./types";

const API_BASE = `${import.meta.env.BASE_URL}api`;

export class ApiError extends Error {
  status: number;
  payload: Record<string, unknown>;

  constructor(status: number, payload: Record<string, unknown>) {
    super(String(payload.error ?? `请求失败（${status}）`));
    this.status = status;
    this.payload = payload;
  }
}

async function json<T>(path: string, init?: RequestInit): Promise<T> {
  const response = await fetch(`${API_BASE}${path}`, {
    ...init,
    headers: init?.body
      ? { "Content-Type": "application/json", ...init.headers }
      : init?.headers,
  });
  const raw = await response.text();
  let payload: Record<string, unknown>;
  try {
    payload = JSON.parse(raw) as Record<string, unknown>;
  } catch {
    throw new ApiError(response.status || 502, { error: raw.slice(0, 180) || "接口返回格式错误" });
  }
  if (!response.ok) throw new ApiError(response.status, payload);
  return payload as T;
}

export function getState(limit?: number): Promise<GridBookState> {
  const query = limit ? `?limit=${Math.max(1, Math.min(10000, Math.floor(limit)))}` : "";
  return json<GridBookState>(`/state${query}`);
}

export function updateCell(
  identity: CollaborationIdentity,
  rowId: number,
  field: EditableField,
  value: string | number | null,
  version: number,
): Promise<Record<string, unknown>> {
  return json(`/cells/${rowId}/${field}`, {
    method: "PATCH",
    body: JSON.stringify({
      client_id: identity.clientId,
      name: identity.name,
      value,
      version,
    }),
  });
}

export function updateCells(
  identity: CollaborationIdentity,
  updates: Array<{
    row_id: number;
    field: EditableField;
    value: string | number | null;
    version: number;
  }>,
): Promise<{ updated: number }> {
  return json("/cells/batch", {
    method: "POST",
    body: JSON.stringify({
      client_id: identity.clientId,
      name: identity.name,
      updates,
    }),
  });
}

export function submitOperations(
  identity: CollaborationIdentity,
  baseRevision: number,
  operations: Array<Record<string, unknown>>,
): Promise<{ revision: number; events?: unknown[]; updated?: number }> {
  return json("/ops/batch", {
    method: "POST",
    body: JSON.stringify({
      client_id: identity.clientId,
      name: identity.name,
      base_revision: baseRevision,
      operations,
    }),
  });
}

export function addRow(
  identity: CollaborationIdentity,
  afterRowId?: number,
): Promise<Record<string, unknown>> {
  return json("/rows", {
    method: "POST",
    body: JSON.stringify({
      client_id: identity.clientId,
      name: identity.name,
      after_row_id: afterRowId,
    }),
  });
}

export function ensureGrid(
  identity: CollaborationIdentity,
  count = 100,
): Promise<Record<string, unknown>> {
  return json("/grid/ensure", {
    method: "POST",
    body: JSON.stringify({
      client_id: identity.clientId,
      name: identity.name,
      count,
    }),
  });
}

export function deleteRows(
  identity: CollaborationIdentity,
  rowIds: number[],
): Promise<Record<string, unknown>> {
  return json("/rows", {
    method: "DELETE",
    body: JSON.stringify({
      client_id: identity.clientId,
      name: identity.name,
      row_ids: rowIds,
    }),
  });
}

export function updateLayout(
  identity: CollaborationIdentity,
  action: string,
  payload: Record<string, unknown>,
): Promise<Record<string, unknown>> {
  return json("/layout", {
    method: "POST",
    body: JSON.stringify({
      client_id: identity.clientId,
      name: identity.name,
      action,
      ...payload,
    }),
  });
}

export function formatCells(
  identity: CollaborationIdentity,
  cells: Array<{ row_id: number; field: EditableField }>,
  format: CellFormat,
): Promise<Record<string, unknown>> {
  return updateLayout(identity, "format", { cells, format });
}

export type BorderAction = "none" | "all" | "outside" | "top" | "bottom" | "left" | "right";

export function updateBorders(
  identity: CollaborationIdentity,
  cells: Array<{ row_id: number; field: EditableField }>,
  action: BorderAction,
  border: CellBorders["top"] = { style: "solid", width: 1, color: "#66717d" },
): Promise<Record<string, unknown>> {
  // Keep the legacy boolean in the payload as a compatibility path for the
  // first-generation backend. Newer servers use border_action + borders for
  // per-edge persistence; older servers ignore border_action but still store
  // the boolean format field.  Without this fallback an old server returns a
  // successful revision with no border state, and its next SSE snapshot rolls
  // back the optimistic border in the browser.
  return updateLayout(identity, "format", {
    cells,
    format: { border: action !== "none" },
    border_action: action,
    border,
  });
}

export function updatePresence(
  identity: CollaborationIdentity,
  cell: { row_id: number; field: EditableField } | null,
): Promise<{ presence: PresenceItem[] }> {
  return json("/presence", {
    method: "POST",
    body: JSON.stringify({
      client_id: identity.clientId,
      name: identity.name,
      cell,
    }),
  });
}

export function eventsUrl(identity: CollaborationIdentity, limit?: number): string {
  const params = new URLSearchParams({
    client_id: identity.clientId,
    name: identity.name,
  });
  if (limit) params.set("limit", String(Math.max(1, Math.min(10000, Math.floor(limit)))));
  return `${API_BASE}/events?${params.toString()}`;
}
