import { promises as fs } from "node:fs";
import path from "node:path";
import { Pool } from "pg";
import type { ExistingShipmentMatch, ShipmentRow } from "./order-types";
import { normalizeText, validateShipmentRow, valueToText } from "./order-utils";

type StoredShipment = ShipmentRow & {
  id: string;
  batchId: string;
  sourceRowNumber: number;
  sourceFileName: string;
  submittedAt: string;
};

type ShipmentListResult = {
  total: number;
  rows: StoredShipment[];
};

const localDataDir = path.join(process.cwd(), ".data");
const localDataFile = path.join(localDataDir, "orders.json");

let pool: Pool | null = null;
let schemaPromise: Promise<void> | null = null;

function hasDatabase(): boolean {
  return Boolean(process.env.POSTGRES_URL || process.env.DATABASE_URL);
}

function getPool(): Pool {
  if (!pool) {
    const connectionString = process.env.POSTGRES_URL || process.env.DATABASE_URL;
    if (!connectionString) {
      throw new Error("DATABASE_UNAVAILABLE");
    }
    pool = new Pool({ connectionString, max: 1 });
  }
  return pool;
}

async function ensureSchema(): Promise<void> {
  if (!hasDatabase()) return;
  if (!schemaPromise) {
    schemaPromise = (async () => {
      const client = await getPool().connect();
      try {
        await client.query(`
          create table if not exists shipments (
            id text primary key,
            external_code text,
            sender_name text not null,
            sender_phone text not null,
            sender_address text not null,
            recipient_name text not null,
            recipient_phone text not null,
            recipient_address text not null,
            weight_kg text not null,
            quantity text not null,
            temp_zone text not null,
            note text,
            batch_id text not null,
            source_row_number integer not null,
            source_file_name text not null,
            submitted_at timestamptz not null
          );
        `);
        await client.query(`
          create unique index if not exists shipments_external_code_unique
          on shipments (external_code)
          where external_code is not null and external_code <> '';
        `);
        await client.query(`
          create index if not exists shipments_submitted_at_idx
          on shipments (submitted_at desc);
        `);
        await client.query(`
          create index if not exists shipments_external_code_idx
          on shipments (external_code);
        `);
        await client.query(`
          create index if not exists shipments_recipient_name_idx
          on shipments (recipient_name);
        `);
      } finally {
        client.release();
      }
    })();
  }
  await schemaPromise;
}

async function ensureLocalFile(): Promise<void> {
  await fs.mkdir(localDataDir, { recursive: true });
  try {
    await fs.access(localDataFile);
  } catch {
    await fs.writeFile(localDataFile, "[]", "utf8");
  }
}

async function readLocalShipments(): Promise<StoredShipment[]> {
  await ensureLocalFile();
  const raw = await fs.readFile(localDataFile, "utf8");
  const parsed = JSON.parse(raw || "[]") as StoredShipment[];
  return Array.isArray(parsed) ? parsed : [];
}

async function writeLocalShipments(rows: StoredShipment[]): Promise<void> {
  await ensureLocalFile();
  await fs.writeFile(localDataFile, JSON.stringify(rows, null, 2), "utf8");
}

export async function listShipments(args: {
  page: number;
  pageSize: number;
  externalCode?: string;
  recipientName?: string;
  submittedFrom?: string;
  submittedTo?: string;
}): Promise<ShipmentListResult> {
  const page = Math.max(args.page, 1);
  const pageSize = Math.max(Math.min(args.pageSize, 100), 1);
  if (hasDatabase()) {
    await ensureSchema();
    const conditions: string[] = [];
    const values: unknown[] = [];
    if (args.externalCode) {
      values.push(`%${args.externalCode.trim()}%`);
      conditions.push(`external_code ilike $${values.length}`);
    }
    if (args.recipientName) {
      values.push(`%${args.recipientName.trim()}%`);
      conditions.push(`recipient_name ilike $${values.length}`);
    }
    if (args.submittedFrom) {
      values.push(new Date(args.submittedFrom).toISOString());
      conditions.push(`submitted_at >= $${values.length}`);
    }
    if (args.submittedTo) {
      values.push(new Date(`${args.submittedTo}T23:59:59.999Z`).toISOString());
      conditions.push(`submitted_at <= $${values.length}`);
    }
    const where = conditions.length ? `where ${conditions.join(" and ")}` : "";
    const countQuery = await getPool().query(`select count(*)::int as total from shipments ${where}`, values);
    const offset = (page - 1) * pageSize;
    const listValues = [...values, pageSize, offset];
    const listQuery = await getPool().query(
      `select * from shipments ${where} order by submitted_at desc limit $${listValues.length - 1} offset $${listValues.length}`,
      listValues
    );
    return {
      total: countQuery.rows[0]?.total ?? 0,
      rows: listQuery.rows.map(dbRowToStoredShipment)
    };
  }
  const rows = await readLocalShipments();
  const filtered = rows.filter((row) => {
    if (args.externalCode && !normalizeText(row.externalCode).includes(normalizeText(args.externalCode))) return false;
    if (args.recipientName && !normalizeText(row.recipientName).includes(normalizeText(args.recipientName))) return false;
    if (args.submittedFrom && new Date(row.submittedAt) < new Date(args.submittedFrom)) return false;
    if (args.submittedTo && new Date(row.submittedAt) > new Date(`${args.submittedTo}T23:59:59.999Z`)) return false;
    return true;
  });
  const offset = (page - 1) * pageSize;
  return {
    total: filtered.length,
    rows: filtered.slice(offset, offset + pageSize)
  };
}

function dbRowToStoredShipment(row: Record<string, unknown>): StoredShipment {
  return {
    id: String(row.id),
    externalCode: valueToText(row.external_code),
    senderName: valueToText(row.sender_name),
    senderPhone: valueToText(row.sender_phone),
    senderAddress: valueToText(row.sender_address),
    recipientName: valueToText(row.recipient_name),
    recipientPhone: valueToText(row.recipient_phone),
    recipientAddress: valueToText(row.recipient_address),
    weightKg: valueToText(row.weight_kg),
    quantity: valueToText(row.quantity),
    tempZone: valueToText(row.temp_zone),
    note: valueToText(row.note),
    batchId: valueToText(row.batch_id),
    sourceRowNumber: Number(row.source_row_number),
    sourceFileName: valueToText(row.source_file_name),
    submittedAt: new Date(valueToText(row.submitted_at)).toISOString()
  };
}

export async function findExistingShipmentsByCodes(codes: string[]): Promise<ExistingShipmentMatch[]> {
  const normalized = Array.from(
    new Set(
      codes
        .map((code) => valueToText(code))
        .map((code) => code.trim())
        .filter(Boolean)
    )
  );
  if (!normalized.length) return [];
  if (hasDatabase()) {
    await ensureSchema();
    const result = await getPool().query(
      `select external_code, source_row_number, batch_id, submitted_at
       from shipments
       where external_code = any($1::text[])`,
      [normalized]
    );
    return result.rows.map((row) => ({
      externalCode: row.external_code as string,
      sourceRowNumber: Number(row.source_row_number),
      batchId: String(row.batch_id),
      submittedAt: new Date(row.submitted_at).toISOString()
    }));
  }
  const rows = await readLocalShipments();
  return rows
    .filter((row) => row.externalCode && normalized.includes(row.externalCode))
    .map((row) => ({
      externalCode: row.externalCode,
      sourceRowNumber: row.sourceRowNumber,
      batchId: row.batchId,
      submittedAt: row.submittedAt
    }));
}

export async function saveShipments(args: {
  rows: ShipmentRow[];
  sourceRowNumbers: number[];
  sourceFileName: string;
}): Promise<{ successCount: number; failureCount: number; inserted: StoredShipment[]; failures: Array<{ index: number; reason: string; sourceRowNumber: number }> }> {
  const failures: Array<{ index: number; reason: string; sourceRowNumber: number }> = [];
  const inserted: StoredShipment[] = [];
  const batchId = crypto.randomUUID();
  const submittedAt = new Date().toISOString();
  if (hasDatabase()) {
    await ensureSchema();
    for (let index = 0; index < args.rows.length; index++) {
      const row = args.rows[index];
      const sourceRowNumber = args.sourceRowNumbers[index] ?? index + 1;
      try {
        const rowIssues = validateShipmentRow(row);
        if (rowIssues.length) {
          failures.push({
            index,
            sourceRowNumber,
            reason: rowIssues.map((issue) => `${issue.fieldLabel}：${issue.message}`).join("；")
          });
          continue;
        }
        if (row.externalCode.trim()) {
          const duplicate = await getPool().query(
            `select source_row_number, batch_id, submitted_at
             from shipments
             where external_code = $1
             limit 1`,
            [row.externalCode]
          );
          if (duplicate.rowCount) {
            const existing = duplicate.rows[0];
            failures.push({
              index,
              reason: `与历史记录第 ${existing.source_row_number} 行重复`,
              sourceRowNumber
            });
            continue;
          }
        }
        const id = crypto.randomUUID();
        await getPool().query(
          `insert into shipments (
            id, external_code, sender_name, sender_phone, sender_address, recipient_name,
            recipient_phone, recipient_address, weight_kg, quantity, temp_zone, note,
            batch_id, source_row_number, source_file_name, submitted_at
          ) values (
            $1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16
          )`,
          [
            id,
            row.externalCode || null,
            row.senderName,
            row.senderPhone,
            row.senderAddress,
            row.recipientName,
            row.recipientPhone,
            row.recipientAddress,
            row.weightKg,
            row.quantity,
            row.tempZone,
            row.note || null,
            batchId,
            sourceRowNumber,
            args.sourceFileName,
            submittedAt
          ]
        );
        inserted.push({
          ...row,
          id,
          batchId,
          sourceRowNumber,
          sourceFileName: args.sourceFileName,
          submittedAt
        });
      } catch (error) {
        failures.push({
          index,
          sourceRowNumber,
          reason: error instanceof Error ? error.message : "保存失败"
        });
      }
    }
    return {
      successCount: inserted.length,
      failureCount: failures.length,
      inserted,
      failures
    };
  }
  const current = await readLocalShipments();
  for (let index = 0; index < args.rows.length; index++) {
    const row = args.rows[index];
    const sourceRowNumber = args.sourceRowNumbers[index] ?? index + 1;
    const rowIssues = validateShipmentRow(row);
    if (rowIssues.length) {
      failures.push({
        index,
        sourceRowNumber,
        reason: rowIssues.map((issue) => `${issue.fieldLabel}：${issue.message}`).join("；")
      });
      continue;
    }
    if (row.externalCode.trim()) {
      const duplicate = current.find((item) => item.externalCode === row.externalCode);
      if (duplicate) {
        failures.push({
          index,
          sourceRowNumber,
          reason: `与历史记录第 ${duplicate.sourceRowNumber} 行重复`
        });
        continue;
      }
    }
    const id = crypto.randomUUID();
    const stored: StoredShipment = {
      ...row,
      id,
      batchId,
      sourceRowNumber,
      sourceFileName: args.sourceFileName,
      submittedAt
    };
    current.push(stored);
    inserted.push(stored);
  }
  await writeLocalShipments(current);
  return {
    successCount: inserted.length,
    failureCount: failures.length,
    inserted,
    failures
  };
}
