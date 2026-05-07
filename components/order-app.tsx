"use client";

import * as XLSX from "xlsx";
import {
  Database,
  Download,
  FileSpreadsheet,
  Loader2,
  Plus,
  RefreshCcw,
  Search,
  Send,
  Settings2,
  ShieldAlert,
  Table2,
  Trash2,
  UploadCloud,
  CheckCircle2,
  Save
} from "lucide-react";
import { useCallback, useEffect, useMemo, useRef, useState, type ChangeEvent, type KeyboardEvent } from "react";
import { exportShipmentsToWorkbook, parseWorkbook, workbookToDownloadFileName } from "../lib/excel";
import type { ExistingShipmentMatch, FieldKey, ImportSession, Mapping, RowIssue, ShipmentRow, ValidatedShipmentRow } from "../lib/order-types";
import { FIELD_META, TEMP_ZONES } from "../lib/order-types";
import {
  buildColumnLabel,
  inferMappingFromLabels,
  mappingHasRequiredColumns,
  mappingToString,
  normalizeShipmentRow,
  stringToMapping,
  valueToText
} from "../lib/order-utils";

type ToastItem = {
  id: string;
  tone: "success" | "error" | "info";
  title: string;
  message: string;
};

type EditableRow = ShipmentRow & {
  rowId: string;
  sourceRowNumber: number;
  rowIndex: number;
};

type HistoryRow = {
  id: string;
  externalCode: string;
  senderName: string;
  senderPhone: string;
  senderAddress: string;
  recipientName: string;
  recipientPhone: string;
  recipientAddress: string;
  weightKg: string;
  quantity: string;
  tempZone: string;
  note: string;
  batchId: string;
  sourceRowNumber: number;
  sourceFileName: string;
  submittedAt: string;
};

type HistoryResponse = {
  total: number;
  rows: HistoryRow[];
};

const APP_TAB_KEY = "excel-order-importer.template-memory";

function getColumnOptions(labels: string[]): Array<{ label: string; index: number }> {
  return labels.map((label, index) => ({ label: label || buildColumnLabel(index), index }));
}

function loadTemplateMemory(): Record<string, string> {
  if (typeof window === "undefined") return {};
  try {
    const raw = window.localStorage.getItem(APP_TAB_KEY);
    return raw ? (JSON.parse(raw) as Record<string, string>) : {};
  } catch {
    return {};
  }
}

function saveTemplateMemory(signature: string, mapping: Mapping) {
  if (typeof window === "undefined") return;
  const next = loadTemplateMemory();
  next[signature] = mappingToString(mapping);
  window.localStorage.setItem(APP_TAB_KEY, JSON.stringify(next));
}

function loadTemplateMapping(signature: string): Mapping | null {
  const memory = loadTemplateMemory();
  const direct = memory[signature];
  if (direct) return stringToMapping(direct);
  return null;
}

function createEmptyRow(rowIndex: number): EditableRow {
  return {
    rowId: crypto.randomUUID(),
    rowIndex,
    sourceRowNumber: rowIndex + 1,
    externalCode: "",
    senderName: "",
    senderPhone: "",
    senderAddress: "",
    recipientName: "",
    recipientPhone: "",
    recipientAddress: "",
    weightKg: "",
    quantity: "",
    tempZone: TEMP_ZONES[0],
    note: ""
  };
}

function formatDateTime(value: string): string {
  if (!value) return "-";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return new Intl.DateTimeFormat("zh-CN", {
    dateStyle: "medium",
    timeStyle: "short"
  }).format(date);
}

function summarizeIssues(issues: RowIssue[]): string {
  return issues.map((issue) => `${issue.fieldLabel}：${issue.message}`).join("；");
}

function buildValidatedRows(
  rows: EditableRow[],
  duplicateMatches: Record<string, ExistingShipmentMatch>,
  serverIssues: Record<string, RowIssue[]>
): {
  rows: ValidatedShipmentRow[];
  issuesByRowId: Record<string, RowIssue[]>;
  allIssues: Array<{ rowId: string; sourceRowNumber: number; text: string }>;
} {
  const issuesByRowId: Record<string, RowIssue[]> = {};
  const firstSeen = new Map<string, EditableRow>();
  const batchDuplicates = new Map<string, EditableRow[]>();

  for (const row of rows) {
    const normalizedCode = row.externalCode.trim();
    if (!normalizedCode) continue;
    const first = firstSeen.get(normalizedCode);
    if (!first) {
      firstSeen.set(normalizedCode, row);
      continue;
    }
    const entries = batchDuplicates.get(normalizedCode) ?? [first];
    if (!batchDuplicates.has(normalizedCode)) {
      batchDuplicates.set(normalizedCode, entries);
    }
    entries.push(row);
  }

  for (const [, groupedRows] of batchDuplicates.entries()) {
    const first = groupedRows[0];
    for (let index = 1; index < groupedRows.length; index++) {
      const row = groupedRows[index];
      const issue = {
        fieldKey: "externalCode" as FieldKey,
        fieldLabel: "外部编码",
        message: `与第 ${first.sourceRowNumber} 行重复`
      };
      const nextIssues = issuesByRowId[row.rowId] ?? [];
      nextIssues.push(issue);
      issuesByRowId[row.rowId] = nextIssues;
      const firstIssue = {
        fieldKey: "externalCode" as FieldKey,
        fieldLabel: "外部编码",
        message: `与第 ${row.sourceRowNumber} 行重复`
      };
      const firstIssues = issuesByRowId[first.rowId] ?? [];
      firstIssues.push(firstIssue);
      issuesByRowId[first.rowId] = firstIssues;
    }
  }

  const validated = rows.map((row) => {
    const duplicate = duplicateMatches[row.externalCode.trim()];
    const baseIssues = [...(issuesByRowId[row.rowId] ?? []), ...(serverIssues[row.rowId] ?? [])];
    if (duplicate && row.externalCode.trim()) {
      baseIssues.push({
        fieldKey: "externalCode",
        fieldLabel: "外部编码",
        message: `与历史记录第 ${duplicate.sourceRowNumber} 行重复`
      });
    }
    issuesByRowId[row.rowId] = baseIssues;
    return {
      ...row,
      ...normalizeShipmentRow(row),
      issues: baseIssues,
      duplicateExternalCode: duplicate?.externalCode,
      duplicateSourceRow: duplicate?.sourceRowNumber
    };
  });

  const allIssues = validated.flatMap((row) =>
    row.issues.map((issue) => ({
      rowId: row.rowId,
      sourceRowNumber: row.sourceRowNumber,
      text: `第 ${row.sourceRowNumber} 行 ${issue.fieldLabel}：${issue.message}`
    }))
  );
  return { rows: validated, issuesByRowId, allIssues };
}

function buildRowFromSession(session: ImportSession, mapping: Mapping, rawRow: string[], rowIndex: number): EditableRow {
  const getCell = (field: FieldKey) => {
    const columnIndex = mapping[field];
    if (!Number.isInteger(columnIndex)) return "";
    return valueToText(rawRow[columnIndex as number]);
  };
  return {
    rowId: `${session.fileName}-${rowIndex}-${crypto.randomUUID()}`,
    rowIndex,
    sourceRowNumber: session.dataStartRowIndex + rowIndex + 1,
    externalCode: getCell("externalCode"),
    senderName: getCell("senderName"),
    senderPhone: getCell("senderPhone"),
    senderAddress: getCell("senderAddress"),
    recipientName: getCell("recipientName"),
    recipientPhone: getCell("recipientPhone"),
    recipientAddress: getCell("recipientAddress"),
    weightKg: getCell("weightKg"),
    quantity: getCell("quantity"),
    tempZone: getCell("tempZone") || TEMP_ZONES[0],
    note: getCell("note")
  };
}

function buildRowsFromSession(session: ImportSession, mapping: Mapping): EditableRow[] {
  return session.rawRows.map((rawRow, index) => buildRowFromSession(session, mapping, rawRow, index));
}

function exportRows(rows: EditableRow[], fileName: string) {
  const workbook = exportShipmentsToWorkbook(rows);
  XLSX.writeFile(workbook, workbookToDownloadFileName(fileName));
}

const EMPTY_HISTORY: HistoryResponse = { total: 0, rows: [] };

export function OrderApp() {
  const [tab, setTab] = useState<"import" | "preview" | "history">("import");
  const [session, setSession] = useState<ImportSession | null>(null);
  const [rows, setRows] = useState<EditableRow[]>([]);
  const [mapping, setMapping] = useState<Mapping>({});
  const [duplicateMatches, setDuplicateMatches] = useState<Record<string, ExistingShipmentMatch>>({});
  const [serverIssues, setServerIssues] = useState<Record<string, RowIssue[]>>({});
  const [busy, setBusy] = useState<"idle" | "parsing" | "submitting" | "loading-history">("idle");
  const [importProgress, setImportProgress] = useState({ percent: 0, current: 0, total: 0, label: "等待上传" });
  const [submitProgress, setSubmitProgress] = useState({ percent: 0, current: 0, total: 0, label: "等待提交" });
  const [dragActive, setDragActive] = useState(false);
  const [history, setHistory] = useState<HistoryResponse>(EMPTY_HISTORY);
  const [historyPage, setHistoryPage] = useState(1);
  const [historyPageSize] = useState(20);
  const [historyFilters, setHistoryFilters] = useState({ externalCode: "", recipientName: "", submittedFrom: "", submittedTo: "" });
  const [toasts, setToasts] = useState<ToastItem[]>([]);
  const [manualMode, setManualMode] = useState(false);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const cellRefs = useRef<Record<string, HTMLInputElement | null>>({});

  const loadHistory = useCallback(
    async (page = historyPage) => {
      setBusy("loading-history");
      try {
        const params = new URLSearchParams({
          page: String(page),
          pageSize: String(historyPageSize),
          externalCode: historyFilters.externalCode,
          recipientName: historyFilters.recipientName,
          submittedFrom: historyFilters.submittedFrom,
          submittedTo: historyFilters.submittedTo
        });
        const response = await fetch(`/api/orders?${params.toString()}`);
        const data = (await response.json()) as HistoryResponse;
        setHistory(data ?? EMPTY_HISTORY);
      } finally {
        setBusy("idle");
      }
    },
    [historyFilters, historyPage, historyPageSize]
  );

  useEffect(() => {
    if (tab !== "history") return;
    void loadHistory(historyPage);
  }, [historyPage, loadHistory, tab]);

  const validated = useMemo(() => buildValidatedRows(rows, duplicateMatches, serverIssues), [rows, duplicateMatches, serverIssues]);
  const totalErrors = validated.allIssues.length;
  const validRows = useMemo(() => validated.rows.filter((row) => row.issues.length === 0), [validated.rows]);
  const canSubmit = rows.length > 0 && totalErrors === 0 && busy !== "submitting";
  const requiredMappingReady = mappingHasRequiredColumns(mapping);

  useEffect(() => {
    if (!rows.length) {
      setDuplicateMatches({});
      return;
    }
    const codes = Array.from(new Set(rows.map((row) => row.externalCode.trim()).filter(Boolean)));
    if (!codes.length) {
      setDuplicateMatches({});
      return;
    }
    const controller = new AbortController();
    const timer = window.setTimeout(async () => {
      try {
        const response = await fetch("/api/orders/existing-codes", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ codes }),
          signal: controller.signal
        });
        if (!response.ok) return;
        const data = (await response.json()) as { matches: ExistingShipmentMatch[] };
        const next = Object.fromEntries(data.matches.map((match) => [match.externalCode, match]));
        setDuplicateMatches(next);
      } catch {
        if (!controller.signal.aborted) return;
      }
    }, 300);
    return () => {
      controller.abort();
      window.clearTimeout(timer);
    };
  }, [rows]);

  useEffect(() => {
    if (typeof window === "undefined") return;
    const saved = loadTemplateMemory();
    if (session?.signature && saved[session.signature]) {
      const remembered = stringToMapping(saved[session.signature]);
      setMapping(remembered);
    }
  }, [session?.signature]);

  useEffect(() => {
    if (!session) return;
    if (manualMode) {
      saveTemplateMemory(session.signature, mapping);
    }
  }, [mapping, manualMode, session]);

  const pushToast = useCallback((tone: ToastItem["tone"], title: string, message: string) => {
    const item: ToastItem = { id: crypto.randomUUID(), tone, title, message };
    setToasts((current) => [item, ...current].slice(0, 3));
    window.setTimeout(() => {
      setToasts((current) => current.filter((toast) => toast.id !== item.id));
    }, 3500);
  }, []);

  const openFileDialog = useCallback(() => {
    fileInputRef.current?.click();
  }, []);

  const parseFile = useCallback(
    async (file: File) => {
      const lower = file.name.toLowerCase();
      if (!lower.endsWith(".xlsx") && !lower.endsWith(".xls")) {
        pushToast("error", "文件格式错误", "请上传 .xlsx 或 .xls 文件。");
        return;
      }
      setBusy("parsing");
      setImportProgress({ percent: 5, current: 0, total: 0, label: "读取文件中" });
      try {
        const buffer = await file.arrayBuffer();
        setImportProgress({ percent: 15, current: 0, total: 0, label: "解析工作簿" });
        const parsed = parseWorkbook(buffer);
        const remembered = loadTemplateMapping(parsed.session.signature);
        const effectiveMapping = remembered ?? parsed.session.mapping;
        const sessionSnapshot: ImportSession = {
          ...parsed.session,
          fileName: file.name,
          mapping: effectiveMapping,
          autoDetected: mappingHasRequiredColumns(effectiveMapping)
        };
        setSession(sessionSnapshot);
        setMapping(effectiveMapping);
        setManualMode(!remembered && !mappingHasRequiredColumns(effectiveMapping));
        setServerIssues({});
        const parsedRows: EditableRow[] = [];
        const total = sessionSnapshot.rawRows.length;
        const chunkSize = total > 1000 ? 80 : 50;
        for (let index = 0; index < total; index += chunkSize) {
          const chunk = sessionSnapshot.rawRows.slice(index, index + chunkSize);
          chunk.forEach((rawRow, offset) => {
            parsedRows.push(buildRowFromSession(sessionSnapshot, effectiveMapping, rawRow, index + offset));
          });
          const current = Math.min(index + chunk.length, total);
          setImportProgress({
            percent: total ? Math.min(95, Math.round((current / total) * 100)) : 95,
            current,
            total,
            label: `映射并校验 ${current}/${total}`
          });
          await new Promise((resolve) => window.setTimeout(resolve, 0));
        }
        if (!parsedRows.length) {
          pushToast("error", "文件为空", "没有找到有效数据行。");
          setBusy("idle");
          return;
        }
        setRows(parsedRows);
        setBusy("idle");
        setImportProgress({
          percent: 100,
          current: parsedRows.length,
          total: parsedRows.length,
          label: `完成导入 ${parsedRows.length} 条`
        });
        pushToast("success", "导入完成", `已识别 ${parsed.session.sheetName}，共 ${parsedRows.length} 条。`);
        setTab("preview");
      } catch (error) {
        const message = error instanceof Error ? error.message : "文件解析失败";
        pushToast("error", "导入失败", message.includes("Sheet") ? message : "文件解析失败，请检查模板或编码。");
      } finally {
        setBusy("idle");
      }
    },
    [pushToast]
  );

  const handleFileInput = useCallback(
    async (event: ChangeEvent<HTMLInputElement>) => {
      const file = event.target.files?.[0];
      if (!file) return;
      await parseFile(file);
      event.target.value = "";
    },
    [parseFile]
  );

  const updateMapping = useCallback((field: FieldKey, columnIndex: number | "") => {
    setMapping((current) => ({
      ...current,
      [field]: columnIndex === "" ? undefined : columnIndex
    }));
    setManualMode(true);
  }, []);

  const applyMapping = useCallback(() => {
    if (!session) return;
    const nextRows = buildRowsFromSession({ ...session, mapping }, mapping);
    setRows(nextRows);
    saveTemplateMemory(session.signature, mapping);
    setManualMode(true);
    pushToast("success", "映射已保存", "系统已记忆本次列映射。");
  }, [mapping, pushToast, session]);

  const autoDetect = useCallback(() => {
    if (!session) return;
    const inferred = inferMappingFromLabels(session.columnLabels);
    setMapping(inferred.mapping);
    setManualMode(false);
    pushToast("info", "已恢复智能识别", `识别到 ${inferred.confidence} 个字段。`);
  }, [pushToast, session]);

  const addRow = useCallback(() => {
    setRows((current) => [...current, createEmptyRow(current.length)]);
    setServerIssues({});
  }, []);

  const deleteRow = useCallback((rowId: string) => {
    setRows((current) => current.filter((row) => row.rowId !== rowId));
    setServerIssues((current) => {
      const next = { ...current };
      delete next[rowId];
      return next;
    });
  }, []);

  const updateCell = useCallback((rowId: string, field: FieldKey, value: string) => {
    setRows((current) => current.map((row) => (row.rowId === rowId ? { ...row, [field]: value } : row)));
    setServerIssues((current) => {
      if (!current[rowId]) return current;
      const next = { ...current };
      delete next[rowId];
      return next;
    });
  }, []);

  const handleCellKeyDown = useCallback(
    (event: KeyboardEvent<HTMLInputElement>, rowIndex: number, fieldIndex: number) => {
      const isTab = event.key === "Tab";
      const isEnter = event.key === "Enter";
      if (!isTab && !isEnter) return;
      event.preventDefault();
      const direction = event.shiftKey ? -1 : 1;
      const nextIndex = rowIndex * FIELD_META.length + fieldIndex + direction;
      const row = Math.floor(nextIndex / FIELD_META.length);
      const field = nextIndex % FIELD_META.length;
      const next = cellRefs.current[`${row}:${field}`];
      next?.focus();
    },
    []
  );

  const handleSubmit = useCallback(async () => {
    if (!canSubmit) {
      pushToast("error", "存在错误", "请先修正所有标红字段后再提交。");
      return;
    }
    setBusy("submitting");
    setSubmitProgress({ percent: 0, current: 0, total: validRows.length, label: "准备提交" });
    try {
      const batchSize = 100;
      const failures: Array<{ rowId: string; message: string }> = [];
      let successCount = 0;
      for (let index = 0; index < validRows.length; index += batchSize) {
        const chunk = validRows.slice(index, index + batchSize);
        const response = await fetch("/api/orders", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            rows: chunk.map((row) => normalizeShipmentRow(row)),
            sourceRowNumbers: chunk.map((row) => row.sourceRowNumber),
            sourceFileName: session?.fileName ?? "unknown.xlsx"
          })
        });
        const result = (await response.json()) as {
          successCount: number;
          failureCount: number;
          failures: Array<{ index: number; reason: string; sourceRowNumber: number }>;
        };
        successCount += result.successCount;
        result.failures.forEach((failure) => {
          const targetRow = chunk[failure.index];
          if (targetRow) {
            failures.push({ rowId: targetRow.rowId, message: failure.reason });
          }
        });
        const current = Math.min(index + chunk.length, validRows.length);
        setSubmitProgress({
          percent: Math.round((current / validRows.length) * 100),
          current,
          total: validRows.length,
          label: `提交中 ${current}/${validRows.length}`
        });
      }
      if (failures.length) {
        pushToast("error", "部分提交失败", `成功 ${successCount} 条，失败 ${failures.length} 条。`);
        setServerIssues((current) => {
          const next = { ...current };
          for (const failure of failures) {
            next[failure.rowId] = [
              {
                fieldKey: "externalCode",
                fieldLabel: "外部编码",
                message: failure.message
              }
            ];
          }
          return next;
        });
      } else {
        pushToast("success", "提交成功", `成功 ${successCount} 条，失败 0 条。`);
        setServerIssues({});
        await loadHistory(1);
        setTab("history");
      }
    } catch (error) {
      pushToast("error", "提交异常", error instanceof Error ? error.message : "提交失败");
    } finally {
      setBusy("idle");
    }
  }, [canSubmit, loadHistory, pushToast, session?.fileName, validRows]);

  const clearAll = useCallback(() => {
    setSession(null);
    setRows([]);
    setMapping({});
    setDuplicateMatches({});
    setServerIssues({});
    setManualMode(false);
    setImportProgress({ percent: 0, current: 0, total: 0, label: "等待上传" });
    setSubmitProgress({ percent: 0, current: 0, total: 0, label: "等待提交" });
    pushToast("info", "已清空", "当前导入数据已经清空。");
  }, [pushToast]);

  const renderedIssues = validated.allIssues;
  const summary = useMemo(
    () => ({
      total: rows.length,
      valid: validRows.length,
      errors: totalErrors,
      duplicate: Object.keys(duplicateMatches).length
    }),
    [duplicateMatches, rows.length, totalErrors, validRows.length]
  );

  return (
    <main className="app-shell">
      <div className="toast-stack">
        {toasts.map((toast) => (
          <div key={toast.id} className={`toast ${toast.tone}`}>
            <strong>{toast.title}</strong>
            <div className="muted" style={{ marginTop: 6 }}>
              {toast.message}
            </div>
          </div>
        ))}
      </div>

      <header className="topbar">
        <div className="title-wrap">
          <h1>万能导入 - 多模板自动导入下单系统</h1>
          <p>
            支持 .xlsx / .xls 拖拽导入、自动识别多模板、手动映射记忆、实时校验、批量提交和历史运单查看。
            适配不同列名、列序、标题行和合并表头结构。
          </p>
        </div>
        <div className="toolbar">
          <button className="btn primary" onClick={openFileDialog} disabled={busy === "parsing"}>
            <UploadCloud size={16} /> 上传 Excel
          </button>
          <button className="btn" onClick={() => setTab("preview")} disabled={!rows.length}>
            <Table2 size={16} /> 预览
          </button>
          <button className="btn" onClick={() => setTab("history")}>
            <Database size={16} /> 运单列表
          </button>
        </div>
      </header>

      <nav className="tab-row">
        <button className={`tab ${tab === "import" ? "active" : ""}`} onClick={() => setTab("import")}>
          <UploadCloud size={16} /> 导入
        </button>
        <button className={`tab ${tab === "preview" ? "active" : ""}`} onClick={() => setTab("preview")} disabled={!rows.length}>
          <Table2 size={16} /> 预览编辑
        </button>
        <button className={`tab ${tab === "history" ? "active" : ""}`} onClick={() => setTab("history")}>
          <Database size={16} /> 已导入运单
        </button>
      </nav>

      <input ref={fileInputRef} type="file" accept=".xlsx,.xls" hidden onChange={handleFileInput} />

      {tab === "import" && (
        <section className="panel">
          <div
            className={`section dropzone ${dragActive ? "drag" : ""}`}
            onDragOver={(event) => {
              event.preventDefault();
              setDragActive(true);
            }}
            onDragLeave={() => setDragActive(false)}
            onDrop={async (event) => {
              event.preventDefault();
              setDragActive(false);
              const file = event.dataTransfer.files?.[0];
              if (file) await parseFile(file);
            }}
            onClick={openFileDialog}
          >
            <div className="section-head" style={{ marginBottom: 0 }}>
              <div>
                <h2>上传和模板识别</h2>
                <div className="sub">拖拽文件到这里，或点击选择文件。系统会自动检测 Sheet、表头和映射关系。</div>
              </div>
              <div className="actions">
                <button className="btn" onClick={(event) => { event.stopPropagation(); autoDetect(); }} disabled={!session}>
                  <RefreshCcw size={16} /> 重新识别
                </button>
                <button className="btn" onClick={(event) => { event.stopPropagation(); applyMapping(); }} disabled={!session}>
                  <Save size={16} /> 应用并记忆
                </button>
              </div>
            </div>
            <div className="grid two">
              <div>
                <div className="summary-grid" style={{ marginBottom: 12 }}>
                  <div className="summary-card">
                    <div className="label">导入状态</div>
                    <div className="value" style={{ fontSize: 18 }}>{busy === "parsing" ? "解析中" : session ? "已识别" : "等待上传"}</div>
                  </div>
                  <div className="summary-card">
                    <div className="label">总行数</div>
                    <div className="value">{summary.total}</div>
                  </div>
                  <div className="summary-card">
                    <div className="label">错误数</div>
                    <div className="value" style={{ color: totalErrors ? "var(--danger)" : "var(--success)" }}>{summary.errors}</div>
                  </div>
                  <div className="summary-card">
                    <div className="label">重复编码</div>
                    <div className="value">{summary.duplicate}</div>
                  </div>
                </div>

                <div className="progress" aria-label="导入进度">
                  <span style={{ width: `${importProgress.percent}%` }} />
                </div>
                <div className="status-line" style={{ marginTop: 8 }}>
                  {importProgress.label} {importProgress.total ? `(${importProgress.current}/${importProgress.total})` : ""}
                </div>

                <div style={{ marginTop: 16 }}>
                  <div className="pill">
                    <FileSpreadsheet size={14} /> {session?.fileName ?? "未选择文件"}
                  </div>
                  <div style={{ height: 8 }} />
                  <div className="pill">
                    <Settings2 size={14} /> {session?.sheetName ?? "未选择 Sheet"}
                  </div>
                  <div style={{ height: 8 }} />
                  <div className="pill">
                    <CheckCircle2 size={14} /> {requiredMappingReady ? "必填字段已覆盖" : "必填字段未完整映射"}
                  </div>
                </div>
              </div>

              <div className="dropzone" style={{ minHeight: 0 }}>
                <div className="section-head" style={{ marginBottom: 0 }}>
                  <div>
                    <h2>手动映射</h2>
                    <div className="sub">自动识别失败时，在这里把 Excel 列与系统字段重新对应。保存后会记忆当前模板结构。</div>
                  </div>
                </div>
                <div className="mapping-grid">
                  {FIELD_META.map((field) => (
                    <div className="mapping-row" key={field.key}>
                      <strong>
                        {field.label} {field.required ? <span className="pill warn">必填</span> : <span className="pill">选填</span>}
                      </strong>
                      <select value={mapping[field.key] ?? ""} onChange={(event) => updateMapping(field.key, event.target.value ? Number(event.target.value) : "")}>
                        <option value="">不导入</option>
                        {getColumnOptions(session?.columnLabels ?? []).map((column) => (
                          <option value={column.index} key={column.index}>
                            {column.label}
                          </option>
                        ))}
                      </select>
                      <small>{session ? `样本：${session.rawRows[0]?.[mapping[field.key] ?? -1] ?? "-"}` : "上传后显示样本"}</small>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>
        </section>
      )}

      {tab === "preview" && (
        <section className="panel">
          <div className="section">
            <div className="section-head">
              <div>
                <h2>预览与在线编辑</h2>
                <div className="sub">
                  点击单元格即可编辑，Tab / 回车切换。错误会一次性在右侧列出，并在对应单元格下方标红。
                </div>
              </div>
              <div className="actions">
                <button className="btn" onClick={addRow}>
                  <Plus size={16} /> 新增空行
                </button>
                <button className="btn" onClick={() => exportRows(rows, session?.fileName ?? "订单预览.xlsx")} disabled={!rows.length}>
                  <Download size={16} /> 导出 Excel
                </button>
                <button className="btn primary" onClick={handleSubmit} disabled={!canSubmit}>
                  {busy === "submitting" ? <Loader2 size={16} className="spin" /> : <Send size={16} />}
                  提交下单
                </button>
                <button className="btn ghost" onClick={clearAll}>
                  <Trash2 size={16} /> 清空
                </button>
              </div>
            </div>

            <div className="summary-grid" style={{ marginBottom: 14 }}>
              <div className="summary-card">
                <div className="label">当前行数</div>
                <div className="value">{rows.length}</div>
              </div>
              <div className="summary-card">
                <div className="label">可提交行</div>
                <div className="value">{validRows.length}</div>
              </div>
              <div className="summary-card">
                <div className="label">错误行</div>
                <div className="value" style={{ color: totalErrors ? "var(--danger)" : "var(--success)" }}>{totalErrors}</div>
              </div>
              <div className="summary-card">
                <div className="label">提交进度</div>
                <div className="value">{submitProgress.percent}%</div>
              </div>
            </div>
            <div className="progress" aria-label="提交进度条">
              <span style={{ width: `${submitProgress.percent}%` }} />
            </div>
            <div className="status-line" style={{ marginTop: 8 }}>
              {submitProgress.label} {submitProgress.total ? `(${submitProgress.current}/${submitProgress.total})` : ""}
            </div>
          </div>

          <div className="section">
            <div className="table-shell">
              <table>
                <thead>
                  <tr>
                    <th style={{ width: 80 }}>行号</th>
                    <th style={{ width: 120 }}>操作</th>
                    {FIELD_META.map((field) => (
                      <th key={field.key}>{field.label}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {rows.map((row, rowIndex) => (
                    <RowEditor
                      key={row.rowId}
                      row={row}
                      rowIndex={rowIndex}
                      issues={validated.issuesByRowId[row.rowId] ?? []}
                      onChange={updateCell}
                      onDelete={deleteRow}
                      onKeyDown={handleCellKeyDown}
                      onFocusRef={(fieldIndex, input) => {
                        cellRefs.current[`${rowIndex}:${fieldIndex}`] = input;
                      }}
                    />
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          <div className="section">
            <div className="section-head">
              <div>
                <h2>错误汇总</h2>
                <div className="sub">所有错误一次性展示，避免逐条修改后反复提交。</div>
              </div>
              <div className="pill bad">
                <ShieldAlert size={14} /> {renderedIssues.length} 条错误
              </div>
            </div>
            <div className="error-panel">
              {renderedIssues.length ? (
                renderedIssues.map((issue) => (
                  <div key={`${issue.rowId}-${issue.text}`} className="error-item">
                    <span>{issue.text}</span>
                    <span>请修正后再提交</span>
                  </div>
                ))
              ) : (
                <div className="pill good">
                  <CheckCircle2 size={14} /> 当前没有校验错误
                </div>
              )}
            </div>
          </div>
        </section>
      )}

      {tab === "history" && (
        <section className="panel">
          <div className="section">
            <div className="section-head">
              <div>
                <h2>已导入运单列表</h2>
                <div className="sub">从数据库读取历史记录，支持外部编码、收件人姓名和提交时间筛选。</div>
              </div>
              <div className="actions">
                <button className="btn" onClick={() => void loadHistory(historyPage)} disabled={busy === "loading-history"}>
                  <RefreshCcw size={16} /> 刷新
                </button>
              </div>
            </div>
            <div className="list-toolbar">
              <div className="field">
                <label>外部编码</label>
                <input value={historyFilters.externalCode} onChange={(event) => setHistoryFilters((current) => ({ ...current, externalCode: event.target.value }))} />
              </div>
              <div className="field">
                <label>收件人姓名</label>
                <input value={historyFilters.recipientName} onChange={(event) => setHistoryFilters((current) => ({ ...current, recipientName: event.target.value }))} />
              </div>
              <div className="field">
                <label>开始日期</label>
                <input type="date" value={historyFilters.submittedFrom} onChange={(event) => setHistoryFilters((current) => ({ ...current, submittedFrom: event.target.value }))} />
              </div>
              <div className="field">
                <label>结束日期</label>
                <input type="date" value={historyFilters.submittedTo} onChange={(event) => setHistoryFilters((current) => ({ ...current, submittedTo: event.target.value }))} />
              </div>
              <button
                className="btn primary"
                onClick={() => {
                  setHistoryPage(1);
                  void loadHistory(1);
                }}
              >
                <Search size={16} /> 搜索
              </button>
            </div>
          </div>

          <div className="section">
            <div className="summary-grid" style={{ marginBottom: 14 }}>
              <div className="summary-card">
                <div className="label">总记录</div>
                <div className="value">{history.total}</div>
              </div>
              <div className="summary-card">
                <div className="label">当前页</div>
                <div className="value">{historyPage}</div>
              </div>
              <div className="summary-card">
                <div className="label">每页数量</div>
                <div className="value">{historyPageSize}</div>
              </div>
              <div className="summary-card">
                <div className="label">加载状态</div>
                <div className="value" style={{ fontSize: 18 }}>{busy === "loading-history" ? "加载中" : "就绪"}</div>
              </div>
            </div>

            <div className="table-shell">
              <table>
                <thead>
                  <tr>
                    <th>提交时间</th>
                    <th>外部编码</th>
                    <th>收件人</th>
                    <th>收件电话</th>
                    <th>重量</th>
                    <th>件数</th>
                    <th>温层</th>
                    <th>来源文件</th>
                  </tr>
                </thead>
                <tbody>
                  {history.rows.map((row) => (
                    <tr key={row.id}>
                      <td>{formatDateTime(row.submittedAt)}</td>
                      <td>{row.externalCode || "-"}</td>
                      <td>{row.recipientName}</td>
                      <td>{row.recipientPhone}</td>
                      <td>{row.weightKg}</td>
                      <td>{row.quantity}</td>
                      <td>{row.tempZone}</td>
                      <td>{row.sourceFileName}</td>
                    </tr>
                  ))}
                  {!history.rows.length && (
                    <tr>
                      <td colSpan={8} style={{ padding: 20 }} className="muted">
                        暂无历史记录
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>

            <div className="actions" style={{ marginTop: 12, justifyContent: "space-between" }}>
              <div className="muted">页面 {historyPage} / {Math.max(1, Math.ceil(history.total / historyPageSize))}</div>
              <div className="actions">
                <button className="btn" onClick={() => setHistoryPage((page) => Math.max(1, page - 1))} disabled={historyPage <= 1}>
                  上一页
                </button>
                <button
                  className="btn"
                  onClick={() => setHistoryPage((page) => (page * historyPageSize < history.total ? page + 1 : page))}
                  disabled={historyPage * historyPageSize >= history.total}
                >
                  下一页
                </button>
              </div>
            </div>
          </div>

          <div className="section">
            <div className="section-head">
              <div>
                <h2>反思题</h2>
                <div className="sub">题目要求中的两个开放题，直接给出我的判断。</div>
              </div>
            </div>
            <div className="grid two">
              <div className="mapping-row">
                <strong>最容易被忽略的 3 个细节点</strong>
                <small>
                  1. 多模板不只是列名不同，还包括标题行、说明行、合并表头和空白行，这会直接影响“表头定位”算法。
                  2. 外部编码的重复判断要同时覆盖批内重复和历史数据重复，否则提交阶段会出现前端看不见、后端才报错的断层。
                  3. 在线编辑后必须全量重校验并一次性列出所有错误，否则导入体验会被“逐条修正”拖慢。
                </small>
              </div>
              <div className="mapping-row">
                <strong>纯人工完成的时间预估</strong>
                <small>
                  如果不借助 AI，我会按 5 到 8 个工作日估算：前 2 天做模板识别和校验引擎，2 天做表格编辑和导出，
                  1 到 2 天接数据库和提交流程，最后 1 天做兼容和部署修正。真正耗时的不是页面，而是各种边界模板和错误场景。
                </small>
              </div>
            </div>
          </div>
        </section>
      )}
    </main>
  );
}

function RowEditor({
    row,
    rowIndex,
    issues,
    onChange,
    onDelete,
    onKeyDown,
    onFocusRef
  }: {
    row: EditableRow;
    rowIndex: number;
    issues: RowIssue[];
    onChange: (rowId: string, field: FieldKey, value: string) => void;
    onDelete: (rowId: string) => void;
    onKeyDown: (event: KeyboardEvent<HTMLInputElement>, rowIndex: number, fieldIndex: number) => void;
    onFocusRef: (fieldIndex: number, input: HTMLInputElement | null) => void;
  }) {
    return (
      <tr className={issues.length ? "error-row" : ""}>
        <td>{row.sourceRowNumber}</td>
        <td>
          <button className="btn" onClick={() => onDelete(row.rowId)} type="button">
            <Trash2 size={14} /> 删除
          </button>
        </td>
        {FIELD_META.map((field, fieldIndex) => {
          const fieldIssues = issues.filter((issue) => issue.fieldKey === field.key);
          return (
            <td key={field.key}>
              <input
                ref={(input) => onFocusRef(fieldIndex, input)}
                className="cell-input"
                value={row[field.key]}
                onChange={(event) => onChange(row.rowId, field.key, event.target.value)}
                onKeyDown={(event) => onKeyDown(event, rowIndex, fieldIndex)}
                placeholder={field.label}
              />
              {fieldIssues.length > 0 && <div className="cell-error" title={summarizeIssues(fieldIssues)}>{summarizeIssues(fieldIssues)}</div>}
            </td>
          );
        })}
      </tr>
    );
  }
