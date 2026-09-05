import { NextResponse } from "next/server";

export interface MatchedRecord {
  id: string;
  sheet1Client: string;
  sheet1Balance: string | number;
  sheet2Code: string;
  sheet2Name: string;
  sheet2Status: string;
  sheet2DueAmount: string | number;
  matchType: "Code Match" | "Name Substring Match" | "Token Match";
  confidence: number;
}

export function normalizeString(str: string): string {
  if (!str) return "";
  return str
    .toLowerCase()
    .replace(/[^\w\s]/gi, " ")
    .replace(/\s+/g, " ")
    .trim();
}

export function extractCode(str: string): string {
  if (!str) return "";
  const match = str.match(/\b([A-Z]{1,4}\d{3,10})\b/i) || str.match(/[-_\s]([A-Z0-9]{4,12})/i);
  return match ? match[1].toUpperCase() : "";
}

export function cleanNameFromClientString(clientStr: string): string {
  if (!clientStr) return "";
  let name = clientStr.trim();
  // Remove trailing code like -SS92239 or - SS46007 or (SS12345) and tags like OLD/NEW
  name = name.replace(/\b(OLD|NEW)\b/gi, "").trim();
  name = name.replace(/[-_\s:(]+\b[A-Z]{1,4}\d{3,10}\b[)]*/gi, "").trim();
  name = name.replace(/[-_\s]+$/, "").trim();
  return name || clientStr;
}

export async function POST(request: Request) {
  try {
    const body = await request.json();
    const { sheet1Rows, headers1, sheet2Rows, headers2 } = body;

    if (!sheet1Rows || !sheet2Rows) {
      return NextResponse.json(
        { success: false, error: "Both Sheet 1 and Sheet 2 data rows are required for matching" },
        { status: 400 }
      );
    }

    // --- Detect Column Indexes for Sheet 1 ---
    let s1ClientIdx = (headers1 || []).findIndex((h: string) => h.toLowerCase().includes("client"));
    let s1BalanceIdx = (headers1 || []).findIndex((h: string) =>
      h.toLowerCase().includes("balance") || h.toLowerCase().includes("amount")
    );

    if (s1ClientIdx === -1) s1ClientIdx = 0;
    if (s1BalanceIdx === -1) {
      s1BalanceIdx = sheet1Rows[0]?.length > 1 ? 1 : 0;
      sheet1Rows.slice(0, 10).forEach((row: any[]) => {
        row.forEach((val: any, idx: number) => {
          if (idx !== s1ClientIdx && (typeof val === "number" || (!isNaN(Number(val)) && val !== ""))) {
            s1BalanceIdx = idx;
          }
        });
      });
    }

    // --- Detect Column Indexes for Sheet 2 ---
    let s2CodeIdx = (headers2 || []).findIndex((h: string) => h.toLowerCase().includes("code"));
    let s2NameIdx = (headers2 || []).findIndex((h: string) =>
      h.toLowerCase().includes("name") || h.toLowerCase().includes("customer")
    );
    let s2StatusIdx = (headers2 || []).findIndex((h: string) => h.toLowerCase().includes("status"));
    let s2DueIdx = (headers2 || []).findIndex((h: string) =>
      h.toLowerCase().includes("due") || h.toLowerCase().includes("amount") || h.toLowerCase().includes("balance")
    );

    const sampleRows2 = sheet2Rows.slice(0, 15);
    const colTypes: Array<{ isCode: number; isNumeric: number; isStatus: number; isName: number }> = [];

    const maxCols2 = Math.max(...sheet2Rows.map((r: any[]) => r.length), 1);
    for (let col = 0; col < maxCols2; col++) {
      let codeScore = 0;
      let numericScore = 0;
      let statusScore = 0;
      let nameScore = 0;

      sampleRows2.forEach((row: any[]) => {
        const val = row[col] !== undefined ? String(row[col]).trim() : "";
        if (!val) return;

        if (val.match(/^[A-Z]{1,4}\d{3,10}$/i)) {
          codeScore += 3;
        } else if (!isNaN(Number(val.replace(/,/g, "")))) {
          numericScore += 2;
        } else if (val.match(/^(active|inactive|pending|valid|invalid)$/i)) {
          statusScore += 3;
        } else if (val.length > 2 && !val.match(/^[A-Z]{1,4}\d{3,10}$/i)) {
          nameScore += 2;
        }
      });

      colTypes.push({ isCode: codeScore, isNumeric: numericScore, isStatus: statusScore, isName: nameScore });
    }

    if (s2CodeIdx === -1) {
      let maxScore = -1;
      colTypes.forEach((t, idx) => {
        if (t.isCode > maxScore) {
          maxScore = t.isCode;
          s2CodeIdx = idx;
        }
      });
      if (s2CodeIdx === -1) s2CodeIdx = 0;
    }

    if (s2DueIdx === -1) {
      let maxScore = -1;
      colTypes.forEach((t, idx) => {
        if (idx !== s2CodeIdx && t.isNumeric > maxScore) {
          maxScore = t.isNumeric;
          s2DueIdx = idx;
        }
      });
      if (s2DueIdx === -1) s2DueIdx = maxCols2 > 1 ? maxCols2 - 1 : 1;
    }

    if (s2NameIdx === -1) {
      let maxScore = -1;
      colTypes.forEach((t, idx) => {
        if (idx !== s2CodeIdx && idx !== s2DueIdx && t.isName > maxScore) {
          maxScore = t.isName;
          s2NameIdx = idx;
        }
      });
      if (s2NameIdx === -1 || s2NameIdx === s2CodeIdx) {
        // Fallback: if Sheet 2 has no separate name column, use column 1 or 0
        s2NameIdx = s2CodeIdx === 0 && maxCols2 > 1 ? 1 : 0;
      }
    }

    if (s2StatusIdx === -1) {
      let maxScore = -1;
      colTypes.forEach((t, idx) => {
        if (idx !== s2CodeIdx && idx !== s2DueIdx && idx !== s2NameIdx && t.isStatus > maxScore) {
          maxScore = t.isStatus;
          s2StatusIdx = idx;
        }
      });
    }

    const matchedRecords: MatchedRecord[] = [];
    const matchedS1Indices = new Set<number>();
    const matchedS2Indices = new Set<number>();

    // --- Perform Matching ---
    sheet1Rows.forEach((row1: any[], idx1: number) => {
      const s1RawClient = row1[s1ClientIdx] !== undefined ? String(row1[s1ClientIdx]).trim() : "";
      const s1Balance = row1[s1BalanceIdx] !== undefined ? row1[s1BalanceIdx] : "";
      if (!s1RawClient) return;

      const s1Code = extractCode(s1RawClient);
      const s1Norm = normalizeString(s1RawClient);

      let bestMatchIdx2 = -1;
      let bestMatchType: "Code Match" | "Name Substring Match" | "Token Match" = "Code Match";
      let bestScore = 0;

      sheet2Rows.forEach((row2: any[], idx2: number) => {
        const s2RawCode = row2[s2CodeIdx] !== undefined ? String(row2[s2CodeIdx]).trim() : "";
        const s2RawName = row2[s2NameIdx] !== undefined ? String(row2[s2NameIdx]).trim() : "";
        const s2Code = extractCode(s2RawCode) || s2RawCode.toUpperCase();
        const s2NormName = normalizeString(s2RawName);

        if (!s2RawName && !s2RawCode) return;

        // Rule A: Exact Code Match
        if (s1Code && s2Code && s1Code === s2Code) {
          bestScore = 100;
          bestMatchIdx2 = idx2;
          bestMatchType = "Code Match";
          return;
        }

        // Rule B: Case-insensitive Substring Matching
        if (s2NormName && s2NormName.length >= 3 && s1Norm.includes(s2NormName)) {
          const score = 90 + Math.min(10, s2NormName.length);
          if (score > bestScore) {
            bestScore = score;
            bestMatchIdx2 = idx2;
            bestMatchType = "Name Substring Match";
          }
        } else if (s1Norm && s1Norm.length >= 3 && s2NormName.includes(s1Norm)) {
          const score = 85;
          if (score > bestScore) {
            bestScore = score;
            bestMatchIdx2 = idx2;
            bestMatchType = "Name Substring Match";
          }
        }

        // Rule C: Token Overlap
        if (bestScore < 80 && s2NormName && s1Norm) {
          const s1Tokens = s1Norm.split(" ").filter((t) => t.length > 2);
          const s2Tokens = s2NormName.split(" ").filter((t) => t.length > 2);
          const commonTokens = s2Tokens.filter((t) => s1Tokens.includes(t));

          if (s2Tokens.length > 0 && commonTokens.length / s2Tokens.length >= 0.6) {
            const score = 75 + (commonTokens.length / s2Tokens.length) * 10;
            if (score > bestScore) {
              bestScore = score;
              bestMatchIdx2 = idx2;
              bestMatchType = "Token Match";
            }
          }
        }
      });

      if (bestMatchIdx2 !== -1 && bestScore >= 70) {
        const row2 = sheet2Rows[bestMatchIdx2];
        matchedS1Indices.add(idx1);
        matchedS2Indices.add(bestMatchIdx2);

        let rawStatus = s2StatusIdx !== -1 && row2[s2StatusIdx] !== undefined ? String(row2[s2StatusIdx]).trim() : "active";
        if (!isNaN(Number(rawStatus)) || rawStatus === String(row2[s2DueIdx])) {
          rawStatus = "active";
        }

        let rawCode = row2[s2CodeIdx] !== undefined ? String(row2[s2CodeIdx]).trim() : "";
        if (!rawCode) rawCode = s1Code;

        let custName = row2[s2NameIdx] !== undefined ? String(row2[s2NameIdx]).trim() : "";
        
        // If customer name is empty, identical to code, or is a code pattern (like SS92239), extract clean name from Sheet 1!
        if (!custName || custName === rawCode || custName.match(/^[A-Z]{1,4}\d{3,10}$/i)) {
          custName = cleanNameFromClientString(s1RawClient);
        }

        matchedRecords.push({
          id: `match-${idx1}-${bestMatchIdx2}`,
          sheet1Client: s1RawClient,
          sheet1Balance: s1Balance,
          sheet2Code: rawCode,
          sheet2Name: custName,
          sheet2Status: rawStatus,
          sheet2DueAmount: row2[s2DueIdx] !== undefined ? row2[s2DueIdx] : 0,
          matchType: bestMatchType,
          confidence: Math.min(100, Math.round(bestScore)),
        });
      }
    });

    const unmatchedSheet1 = sheet1Rows
      .filter((_: any, idx: number) => !matchedS1Indices.has(idx))
      .map((row: any[]) => ({
        client: row[s1ClientIdx] !== undefined ? String(row[s1ClientIdx]) : "",
        balance: row[s1BalanceIdx] !== undefined ? row[s1BalanceIdx] : "",
        rawRow: row,
      }));

    const unmatchedSheet2 = sheet2Rows
      .filter((_: any, idx: number) => !matchedS2Indices.has(idx))
      .map((row: any[]) => ({
        code: row[s2CodeIdx] !== undefined ? String(row[s2CodeIdx]) : "",
        name: row[s2NameIdx] !== undefined ? String(row[s2NameIdx]) : "",
        status: s2StatusIdx !== -1 && row[s2StatusIdx] !== undefined ? String(row[s2StatusIdx]) : "active",
        dueAmount: row[s2DueIdx] !== undefined ? row[s2DueIdx] : "",
        rawRow: row,
      }));

    return NextResponse.json({
      success: true,
      stats: {
        totalSheet1Rows: sheet1Rows.length,
        totalSheet2Rows: sheet2Rows.length,
        matchedCount: matchedRecords.length,
        unmatchedSheet1Count: unmatchedSheet1.length,
        unmatchedSheet2Count: unmatchedSheet2.length,
        matchRate: Math.round((matchedRecords.length / (sheet1Rows.length || 1)) * 100),
      },
      matchedRecords,
      unmatchedSheet1,
      unmatchedSheet2,
    });
  } catch (error: any) {
    return NextResponse.json(
      { success: false, error: error?.message || "Failed to perform sheet matching" },
      { status: 500 }
    );
  }
}
