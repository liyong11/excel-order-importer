import { NextRequest, NextResponse } from "next/server";
import { listShipments, saveShipments } from "../../../lib/storage";
import { normalizeShipmentRow } from "../../../lib/order-utils";
import type { ShipmentRow } from "../../../lib/order-types";

export const runtime = "nodejs";

export async function GET(request: NextRequest) {
  const url = new URL(request.url);
  const page = Number(url.searchParams.get("page") ?? "1");
  const pageSize = Number(url.searchParams.get("pageSize") ?? "20");
  const externalCode = url.searchParams.get("externalCode") ?? "";
  const recipientName = url.searchParams.get("recipientName") ?? "";
  const submittedFrom = url.searchParams.get("submittedFrom") ?? "";
  const submittedTo = url.searchParams.get("submittedTo") ?? "";
  const result = await listShipments({ page, pageSize, externalCode, recipientName, submittedFrom, submittedTo });
  return NextResponse.json(result);
}

export async function POST(request: NextRequest) {
  const body = (await request.json().catch(() => null)) as
    | { rows?: ShipmentRow[]; sourceRowNumbers?: number[]; sourceFileName?: string }
    | null;
  if (!body?.rows?.length || !body?.sourceRowNumbers?.length) {
    return NextResponse.json({ error: "请求体无效" }, { status: 400 });
  }
  const normalized = body.rows.map((row) => normalizeShipmentRow(row));
  const result = await saveShipments({
    rows: normalized,
    sourceRowNumbers: body.sourceRowNumbers,
    sourceFileName: body.sourceFileName ?? "unknown.xlsx"
  });
  return NextResponse.json(result);
}
