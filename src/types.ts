export type EditableField =
  | "account_name"
  | "teacher"
  | "channel"
  | "publish_count"
  | "total_followers"
  | "short_graphic_leads"
  | "live_leads"
  | "openings"
  | "remark"
  | `custom_${number}`;

export interface GridBookRow {
  id: number;
  account_name: string;
  teacher: string;
  channel: string;
  publish_count: number | null;
  total_followers: number | null;
  short_graphic_leads: number | null;
  live_leads: number | null;
  openings: number | null;
  remark: string;
  [key: string]: unknown;
}

export interface CellMeta {
  version: number;
  updated_by: string;
  updated_at: string;
  client_id: string;
}

export interface BorderSide {
  style: "solid";
  width: 1;
  color: string;
}

export interface CellBorders {
  top?: BorderSide | null;
  right?: BorderSide | null;
  bottom?: BorderSide | null;
  left?: BorderSide | null;
}

export interface ActivityItem {
  id: number;
  kind: "cell" | "row";
  row_id: number;
  field: EditableField;
  field_label: string;
  account_name: string;
  value: string | number | null;
  name: string;
  client_id: string;
  updated_at: string;
}

export interface PresenceItem {
  client_id: string;
  name: string;
  cell: { row_id: number; field: EditableField } | null;
  last_seen: string;
}

export interface CellFormat {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  border?: boolean;
  borders?: CellBorders;
  wrap?: boolean;
  align?: "left" | "center" | "right" | "justify";
  vertical_align?: "top" | "middle" | "bottom";
  fill?: string;
  font_color?: string;
  font_family?: string;
  font_size?: number;
  number_format?: "general" | "number" | "date" | "percentage" | "text";
}

export interface CellMerge {
  id: string;
  start_row_id: number;
  end_row_id: number;
  start_field: EditableField;
  end_field: EditableField;
}

export interface GridBookLayout {
  column_widths: Partial<Record<EditableField, number>>;
  row_heights: Record<string, number>;
  merges: CellMerge[];
  formats: Record<string, CellFormat>;
}

export interface GridBookState {
  schema_version: number;
  revision: number;
  next_row_id: number;
  period: { start: string; end: string };
  rows: GridBookRow[];
  cells: Record<string, CellMeta>;
  layout: GridBookLayout;
  activity: ActivityItem[];
  presence: PresenceItem[];
}

export interface CollaborationIdentity {
  clientId: string;
  name: string;
}

export interface CellEvent {
  conflict: false;
  revision: number;
  row_id: number;
  field: EditableField;
  value: string | number | null;
  meta: CellMeta;
  activity: ActivityItem;
}

export interface RowEvent {
  revision: number;
  row: GridBookRow;
  meta: CellMeta;
  activity: ActivityItem;
}

export interface LayoutEvent {
  revision: number;
  layout: GridBookLayout;
  name: string;
  client_id: string;
}
