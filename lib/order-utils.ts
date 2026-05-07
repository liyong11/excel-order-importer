import { FIELD_META, TEMP_ZONES, type FieldKey, type Mapping, type ShipmentRow, type ValidatedShipmentRow, type RowIssue } from "./order-types";

const DIGITS_ONLY = /\d+/g;

export function normalizeText(value: unknown): string {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/[\s\u3000]+/g, "")
    .replace(/[()（）【】\[\]{}.,，。/\\|_-]/g, "");
}

export function valueToText(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "number" && Number.isFinite(value)) {
    return Number.isInteger(value) ? String(value) : String(value).replace(/\.0+$/, "");
  }
  return String(value).trim();
}

export function emptyRow(row: string[]): boolean {
  return row.every((cell) => !valueToText(cell));
}

export function buildColumnLabel(index: number): string {
  let n = index + 1;
  let label = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    label = String.fromCharCode(65 + rem) + label;
    n = Math.floor((n - 1) / 26);
  }
  return `列 ${label}`;
}

export function getFieldLabel(fieldKey: FieldKey): string {
  return FIELD_META.find((field) => field.key === fieldKey)?.label ?? fieldKey;
}

export const FIELD_ALIASES: Record<FieldKey, string[]> = {
  externalCode: ["外部编码", "外部订单号", "客户单号", "订单号", "refcode", "ref", "refcode", "externalcode", "外部单号"],
  senderName: ["发件人姓名", "发件人", "寄件人姓名", "寄件人", "sender", "shipper", "from"],
  senderPhone: ["发件人电话", "发件电话", "寄件人电话", "寄件电话", "sendertel", "senderphone", "shippertel"],
  senderAddress: ["发件人地址", "发货地址", "寄件人地址", "寄件地址", "senderaddress", "shipperaddress"],
  recipientName: ["收件人姓名", "收件人", "收货人姓名", "收货人", "收方", "receiver", "recipient", "consignee"],
  recipientPhone: ["收件人电话", "收货电话", "收方电话", "收货人电话", "receivertel", "recipientphone", "consigneetel"],
  recipientAddress: ["收件人地址", "收货地址", "收方地址", "receiveraddress", "recipientaddress", "consigneeaddress"],
  weightKg: ["重量kg", "重量", "weightkg", "weight", "grossweight", "货重"],
  quantity: ["件数", "数量", "qty", "quantity", "包裹数量", "箱数"],
  tempZone: ["温层", "温度要求", "温度", "tempzone", "温控", "温区"],
  note: ["备注", "附言", "说明", "note", "remarks", "memo"]
};

export function scoreHeaderCell(fieldKey: FieldKey, cell: string): number {
  const normalized = normalizeText(cell);
  if (!normalized) return 0;
  const aliases = FIELD_ALIASES[fieldKey];
  let best = 0;
  for (const alias of aliases) {
    const aliasNorm = normalizeText(alias);
    if (!aliasNorm) continue;
    if (normalized === aliasNorm) {
      best = Math.max(best, 100);
      continue;
    }
    if (normalized.includes(aliasNorm) || aliasNorm.includes(normalized)) {
      best = Math.max(best, 85);
      continue;
    }
    const overlap = tokenOverlap(normalized, aliasNorm);
    if (overlap > 0) {
      best = Math.max(best, 50 + overlap * 15);
    }
  }
  return best;
}

function tokenOverlap(a: string, b: string): number {
  const aTokens = new Set(splitTokens(a));
  const bTokens = new Set(splitTokens(b));
  let hits = 0;
  for (const token of aTokens) {
    if (bTokens.has(token)) hits++;
  }
  return hits / Math.max(aTokens.size, bTokens.size, 1);
}

function splitTokens(value: string): string[] {
  return value
    .split(/[^a-z0-9\u4e00-\u9fa5]+/i)
    .map((part) => part.trim())
    .filter(Boolean);
}

export function inferMappingFromLabels(labels: string[]): { mapping: Mapping; confidence: number } {
  const mapping: Mapping = {};
  let confidence = 0;
  for (const field of FIELD_META) {
    let bestColumn = -1;
    let bestScore = 0;
    labels.forEach((label, index) => {
      const score = scoreHeaderCell(field.key, label);
      if (score > bestScore) {
        bestScore = score;
        bestColumn = index;
      }
    });
    if (bestColumn >= 0 && bestScore >= 45) {
      mapping[field.key] = bestColumn;
      confidence += 1;
    }
  }
  return { mapping, confidence };
}

export function signatureFromLabels(labels: string[]): string {
  const normalized = labels
    .map((label) => normalizeText(label))
    .filter(Boolean);
  return `${labels.length}:${normalized.join("|")}`;
}

export function mappingToString(mapping: Mapping): string {
  return FIELD_META.map((field) => `${field.key}:${mapping[field.key] ?? ""}`).join(";");
}

export function stringToMapping(value: string): Mapping {
  const mapping: Mapping = {};
  for (const chunk of value.split(";")) {
    const [key, raw] = chunk.split(":");
    if (!key || raw === undefined) continue;
    const index = Number(raw);
    if (Number.isInteger(index) && index >= 0) {
      mapping[key as FieldKey] = index;
    }
  }
  return mapping;
}

export function buildEmptyRow(columnCount = FIELD_META.length): string[] {
  return Array.from({ length: columnCount }, () => "");
}

export function mappingHasRequiredColumns(mapping: Mapping): boolean {
  return FIELD_META.filter((field) => field.required).every((field) => Number.isInteger(mapping[field.key]));
}

export function isValidPhone(value: string): boolean {
  const trimmed = value.trim();
  if (!trimmed) return false;
  const digits = trimmed.match(DIGITS_ONLY)?.join("") ?? "";
  if (digits.length < 7) return false;
  return /^[+]?[\d\s-()（）]+$/.test(trimmed);
}

export function validateShipmentRow(row: ShipmentRow): RowIssue[] {
  const issues: RowIssue[] = [];
  const requiredFields = FIELD_META.filter((field) => field.required);
  for (const field of requiredFields) {
    const value = row[field.key].trim();
    if (!value) {
      issues.push({
        fieldKey: field.key,
        fieldLabel: field.label,
        message: `${field.label}不能为空`
      });
    }
  }
  if (row.senderPhone.trim() && !isValidPhone(row.senderPhone)) {
    issues.push({ fieldKey: "senderPhone", fieldLabel: getFieldLabel("senderPhone"), message: "格式错误" });
  }
  if (row.recipientPhone.trim() && !isValidPhone(row.recipientPhone)) {
    issues.push({ fieldKey: "recipientPhone", fieldLabel: getFieldLabel("recipientPhone"), message: "格式错误" });
  }
  const weight = Number(row.weightKg);
  if (!row.weightKg.trim() || !Number.isFinite(weight) || weight <= 0) {
    issues.push({ fieldKey: "weightKg", fieldLabel: getFieldLabel("weightKg"), message: "必须为正数" });
  }
  const quantity = Number(row.quantity);
  if (!row.quantity.trim() || !Number.isInteger(quantity) || quantity <= 0) {
    issues.push({ fieldKey: "quantity", fieldLabel: getFieldLabel("quantity"), message: "必须为正整数" });
  }
  if (!TEMP_ZONES.includes(row.tempZone as (typeof TEMP_ZONES)[number])) {
    issues.push({ fieldKey: "tempZone", fieldLabel: getFieldLabel("tempZone"), message: "必须是 常温 / 冷藏 / 冷冻" });
  }
  return issues;
}

export function normalizeShipmentRow(row: Partial<ShipmentRow>): ShipmentRow {
  return {
    externalCode: valueToText(row.externalCode),
    senderName: valueToText(row.senderName),
    senderPhone: valueToText(row.senderPhone),
    senderAddress: valueToText(row.senderAddress),
    recipientName: valueToText(row.recipientName),
    recipientPhone: valueToText(row.recipientPhone),
    recipientAddress: valueToText(row.recipientAddress),
    weightKg: valueToText(row.weightKg),
    quantity: valueToText(row.quantity),
    tempZone: valueToText(row.tempZone),
    note: valueToText(row.note)
  };
}

export function rowToCsvLikeText(row: ShipmentRow): string {
  return FIELD_META.map((field) => row[field.key] ?? "").join("\t");
}

export function mergeIssues(...issueGroups: RowIssue[][]): RowIssue[] {
  return issueGroups.flat().filter((issue, index, all) => {
    const key = `${issue.fieldKey ?? ""}:${issue.fieldLabel}:${issue.message}`;
    return all.findIndex((item) => `${item.fieldKey ?? ""}:${item.fieldLabel}:${item.message}` === key) === index;
  });
}

export function ensureFieldOrder(mapping: Mapping): FieldKey[] {
  return FIELD_META.map((field) => field.key).filter((key) => Number.isInteger(mapping[key]));
}
