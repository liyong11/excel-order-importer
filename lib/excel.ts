import * as XLSX from "xlsx";
import type { ImportSession, Mapping, ShipmentRow, ValidatedShipmentRow } from "./order-types";
import { FIELD_META, EXPORT_HEADERS } from "./order-types";
import {
  buildColumnLabel,
  emptyRow,
  inferMappingFromLabels,
  mappingHasRequiredColumns,
  scoreHeaderCell,
  signatureFromLabels,
  validateShipmentRow,
  valueToText
} from "./order-utils";

export type ParsedWorkbook = {
  session: ImportSession;
  candidateSheets: Array<{ sheetName: string; signature: string; score: number; headerRowIndex: number; dataStartRowIndex: number; columnLabels: string[] }>;
};

export function readWorkbookFromBuffer(buffer: ArrayBuffer): XLSX.WorkBook {
  return XLSX.read(buffer, { type: "array", cellDates: false, cellNF: false, cellStyles: false });
}

export function detectBestSheet(workbook: XLSX.WorkBook): ParsedWorkbook["candidateSheets"][number] | null {
  const candidateSheets = workbook.SheetNames.map((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json<string[]>(worksheet, { header: 1, raw: false, defval: "", blankrows: true }) as string[][];
    const scanLimit = Math.min(rows.length, 15);
    let best = {
      score: 0,
      headerRowIndex: 0,
      dataStartRowIndex: 0,
      columnLabels: rows[0]?.map((cell, index) => valueToText(cell) || buildColumnLabel(index)) ?? []
    };
    for (let rowIndex = 0; rowIndex < scanLimit; rowIndex++) {
      const row = rows[rowIndex] ?? [];
      const nonEmpty = row.filter((cell) => valueToText(cell)).length;
      if (!nonEmpty) continue;
      let score = 0;
      const labels = row.map((cell) => valueToText(cell));
      for (const field of FIELD_META) {
        let columnBest = 0;
        for (const cell of labels) {
          columnBest = Math.max(columnBest, scoreHeaderCell(field.key, cell));
        }
        score += columnBest;
      }
      if (row.some((cell) => /说明|备注|请勿|必填|选填/.test(valueToText(cell)))) {
        score *= 0.45;
      }
      score += Math.min(nonEmpty, 15) * 3;
      if (score > best.score) {
        best = {
          score,
          headerRowIndex: rowIndex,
          dataStartRowIndex: rowIndex + 1,
          columnLabels: row.map((cell, index) => valueToText(cell) || buildColumnLabel(index))
        };
      }
    }
    const labelRow = best.columnLabels.length ? best.columnLabels : rows[0]?.map((cell, index) => valueToText(cell) || buildColumnLabel(index)) ?? [];
    return {
      sheetName,
      score: best.score,
      headerRowIndex: best.headerRowIndex,
      dataStartRowIndex: best.dataStartRowIndex,
      signature: signatureFromLabels(labelRow),
      columnLabels: labelRow
    };
  });
  return candidateSheets.sort((a, b) => b.score - a.score)[0] ?? null;
}

export function parseWorkbook(buffer: ArrayBuffer, rememberedMapping?: Mapping): ParsedWorkbook {
  const workbook = readWorkbookFromBuffer(buffer);
  const bestSheet = detectBestSheet(workbook);
  if (!bestSheet) {
    throw new Error("没有找到可用的 Sheet，请检查文件内容。");
  }
  const worksheet = workbook.Sheets[bestSheet.sheetName];
  const rows = XLSX.utils.sheet_to_json<string[]>(worksheet, { header: 1, raw: false, defval: "", blankrows: true }) as string[][];
  const normalizedRows = rows.map((row) => row.map((cell) => valueToText(cell)));
  const candidate = normalizeRawSession(bestSheet.sheetName, bestSheet.headerRowIndex, bestSheet.dataStartRowIndex, normalizedRows, bestSheet.columnLabels);
  const auto = inferMappingFromLabels(candidate.columnLabels);
  const mapping = rememberedMapping && Object.keys(rememberedMapping).length ? rememberedMapping : auto.mapping;
  return {
    session: {
      fileName: "",
      sheetName: bestSheet.sheetName,
      headerRowIndex: bestSheet.headerRowIndex,
      dataStartRowIndex: bestSheet.dataStartRowIndex,
      columnLabels: candidate.columnLabels,
      rawRows: candidate.rawRows,
      signature: candidate.signature,
      mapping,
      autoDetected: mappingHasRequiredColumns(mapping)
    },
    candidateSheets: workbook.SheetNames.map((sheetName) => {
      const worksheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json<string[]>(worksheet, { header: 1, raw: false, defval: "", blankrows: true }) as string[][];
      const info = detectBestSheet({ ...workbook, SheetNames: [sheetName] });
      return {
        sheetName,
        score: info?.score ?? 0,
        headerRowIndex: info?.headerRowIndex ?? 0,
        dataStartRowIndex: info?.dataStartRowIndex ?? 0,
        signature: info?.signature ?? signatureFromLabels(rows[0]?.map((cell, index) => valueToText(cell) || buildColumnLabel(index)) ?? []),
        columnLabels: info?.columnLabels ?? rows[0]?.map((cell, index) => valueToText(cell) || buildColumnLabel(index)) ?? []
      };
    })
  };
}

function normalizeRawSession(
  sheetName: string,
  headerRowIndex: number,
  dataStartRowIndex: number,
  rows: string[][],
  columnLabels: string[]
): ImportSession {
  const width = Math.max(FIELD_META.length, columnLabels.length, ...rows.map((row) => row.length));
  const rawRows = rows
    .slice(dataStartRowIndex)
    .filter((row) => !emptyRow(row))
    .map((row) => Array.from({ length: width }, (_, index) => valueToText(row[index])));
  const signature = signatureFromLabels(columnLabels);
  return {
    fileName: "",
    sheetName,
    headerRowIndex,
    dataStartRowIndex,
    columnLabels,
    rawRows,
    signature,
    mapping: inferMappingFromLabels(columnLabels).mapping,
    autoDetected: true
  };
}

export function mapRowsToShipments(session: ImportSession, mapping: Mapping): { rows: ValidatedShipmentRow[]; issues: string[] } {
  const rows: ValidatedShipmentRow[] = [];
  const effectiveMapping = { ...session.mapping, ...mapping };
  for (let index = 0; index < session.rawRows.length; index++) {
    const sourceRowNumber = session.dataStartRowIndex + 1 + index;
    const raw = session.rawRows[index];
    const row: ShipmentRow = {
      externalCode: getCell(raw, effectiveMapping.externalCode),
      senderName: getCell(raw, effectiveMapping.senderName),
      senderPhone: getCell(raw, effectiveMapping.senderPhone),
      senderAddress: getCell(raw, effectiveMapping.senderAddress),
      recipientName: getCell(raw, effectiveMapping.recipientName),
      recipientPhone: getCell(raw, effectiveMapping.recipientPhone),
      recipientAddress: getCell(raw, effectiveMapping.recipientAddress),
      weightKg: getCell(raw, effectiveMapping.weightKg),
      quantity: getCell(raw, effectiveMapping.quantity),
      tempZone: getCell(raw, effectiveMapping.tempZone),
      note: getCell(raw, effectiveMapping.note)
    };
    const issues = validateShipmentRow(row);
    rows.push({
      ...row,
      rowId: `${sourceRowNumber}-${index}`,
      sourceRowNumber,
      rowIndex: index,
      issues
    });
  }
  return { rows, issues: [] };
}

function getCell(row: string[], columnIndex?: number): string {
  if (!Number.isInteger(columnIndex)) return "";
  return valueToText(row[columnIndex as number]);
}

export function exportShipmentsToWorkbook(rows: ShipmentRow[]): XLSX.WorkBook {
  const worksheet = XLSX.utils.json_to_sheet(
    rows.map((row) =>
      FIELD_META.reduce<Record<string, string>>((acc, field) => {
        acc[field.label] = row[field.key];
        return acc;
      }, {})
    ),
    { header: EXPORT_HEADERS }
  );
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "订单预览");
  return workbook;
}

export function workbookToDownloadFileName(baseName: string): string {
  const clean = baseName.replace(/\.(xlsx|xls)$/i, "").trim() || "订单预览";
  return `${clean}-已整理.xlsx`;
}

export function extractColumnPreview(session: ImportSession, mapping: Mapping): Array<{ fieldKey: string; fieldLabel: string; columnIndex?: number; sample: string }> {
  return FIELD_META.map((field) => {
    const columnIndex = mapping[field.key];
    return {
      fieldKey: field.key,
      fieldLabel: field.label,
      columnIndex,
      sample: Number.isInteger(columnIndex) ? session.rawRows[0]?.[columnIndex as number] ?? "" : ""
    };
  });
}

export function headerPreviewFromSession(session: ImportSession): string[] {
  return session.columnLabels.length ? session.columnLabels : Array.from({ length: 11 }, (_, index) => buildColumnLabel(index));
}
