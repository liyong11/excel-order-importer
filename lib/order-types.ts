export const TEMP_ZONES = ["常温", "冷藏", "冷冻"] as const;

export type TempZone = (typeof TEMP_ZONES)[number];

export type FieldKey =
  | "externalCode"
  | "senderName"
  | "senderPhone"
  | "senderAddress"
  | "recipientName"
  | "recipientPhone"
  | "recipientAddress"
  | "weightKg"
  | "quantity"
  | "tempZone"
  | "note";

export type FieldMeta = {
  key: FieldKey;
  label: string;
  required: boolean;
};

export const FIELD_META: FieldMeta[] = [
  { key: "externalCode", label: "外部编码", required: false },
  { key: "senderName", label: "发件人姓名", required: true },
  { key: "senderPhone", label: "发件人电话", required: true },
  { key: "senderAddress", label: "发件人地址", required: true },
  { key: "recipientName", label: "收件人姓名", required: true },
  { key: "recipientPhone", label: "收件人电话", required: true },
  { key: "recipientAddress", label: "收件人地址", required: true },
  { key: "weightKg", label: "重量 (kg)", required: true },
  { key: "quantity", label: "件数", required: true },
  { key: "tempZone", label: "温层", required: true },
  { key: "note", label: "备注", required: false }
];

export type ShipmentRow = Record<FieldKey, string>;

export type RowIssue = {
  fieldKey?: FieldKey;
  fieldLabel: string;
  message: string;
};

export type ValidatedShipmentRow = ShipmentRow & {
  rowId: string;
  sourceRowNumber: number;
  rowIndex: number;
  issues: RowIssue[];
  duplicateExternalCode?: string;
  duplicateSourceRow?: number;
};

export type Mapping = Partial<Record<FieldKey, number>>;

export type ExistingShipmentMatch = {
  externalCode: string;
  sourceRowNumber: number;
  batchId: string;
  submittedAt: string;
};

export type ImportSession = {
  fileName: string;
  sheetName: string;
  headerRowIndex: number;
  dataStartRowIndex: number;
  columnLabels: string[];
  rawRows: string[][];
  signature: string;
  mapping: Mapping;
  autoDetected: boolean;
};

export const EXPORT_HEADERS = FIELD_META.map((field) => field.label);
