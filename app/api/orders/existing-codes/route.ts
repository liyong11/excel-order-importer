import { NextRequest, NextResponse } from "next/server";
import { findExistingShipmentsByCodes } from "../../../../lib/storage";

export const runtime = "nodejs";

export async function POST(request: NextRequest) {
  const body = (await request.json().catch(() => null)) as { codes?: string[] } | null;
  if (!body?.codes) {
    return NextResponse.json({ matches: [] });
  }
  const matches = await findExistingShipmentsByCodes(body.codes);
  return NextResponse.json({ matches });
}
