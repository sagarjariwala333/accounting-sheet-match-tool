import { NextResponse } from "next/server";
import * as XLSX from "xlsx";

export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    const file = formData.get("file") as File | null;

    if (!file) {
      return NextResponse.json(
        { success: false, error: "No Excel file provided" },
        { status: 400 }
      );
    }

    const fileName = file.name;
    if (!fileName.match(/\.(xlsx|xls|csv|txt)$/i)) {
      return NextResponse.json(
        { success: false, error: "Invalid file format. Please upload an Excel (.xlsx, .xls) or CSV file." },
        { status: 400 }
      );
    }

    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    let workbook: XLSX.WorkBook;
    try {
      // Primary: read buffer
      workbook = XLSX.read(buffer, { type: "buffer", cellDates: true, raw: false });
    } catch (err1) {
      try {
        // Fallback 1: read as text string (handles HTML tables & CSV masquerading as .xls)
        const textContent = buffer.toString("utf-8");
        workbook = XLSX.read(textContent, { type: "string", raw: false });
      } catch (err2) {
        // Fallback 2: read binary string
        const binaryStr = buffer.toString("binary");
        workbook = XLSX.read(binaryStr, { type: "binary", raw: false });
      }
    }

    if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
      return NextResponse.json(
        { success: false, error: "Unable to parse Excel workbook structure." },
        { status: 400 }
      );
    }

    const sheetsData: Record<string, any[][]> = {};
    const sheetsMeta: Record<string, { headers: string[]; rows: any[][]; totalRows: number; totalCols: number }> = {};

    workbook.SheetNames.forEach((sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      if (!sheet) return;

      // Extract rows with blankrows: false
      let rawJson: any[][] = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: "",
        blankrows: false,
      });

      // Filter out completely empty rows
      rawJson = rawJson.filter((row) =>
        Array.isArray(row) && row.some((cell) => cell !== undefined && cell !== null && String(cell).trim() !== "")
      );

      if (!rawJson || rawJson.length === 0) {
        sheetsData[sheetName] = [];
        sheetsMeta[sheetName] = { headers: [], rows: [], totalRows: 0, totalCols: 0 };
        return;
      }

      // Calculate max columns
      let maxCols = 0;
      rawJson.forEach((row) => {
        if (Array.isArray(row) && row.length > maxCols) maxCols = row.length;
      });

      // Determine if Row 0 looks like a Header
      // (If first row cells are strings and there's more than 1 row, treat row 0 as header)
      let headers: string[] = [];
      let dataRows: any[][] = [];

      const firstRow = rawJson[0] || [];
      const hasHeaderRow = rawJson.length > 1 && firstRow.some((cell) => typeof cell === "string" && isNaN(Number(cell)));

      if (hasHeaderRow) {
        headers = firstRow.map((h, i) =>
          h !== undefined && h !== null && String(h).trim() !== ""
            ? String(h).trim()
            : `Column ${i + 1}`
        );
        dataRows = rawJson.slice(1);
      } else {
        // Generate generic headers and keep all rows as data
        headers = Array.from({ length: maxCols }, (_, i) => `Column ${i + 1}`);
        dataRows = rawJson;
      }

      // Pad headers up to maxCols
      while (headers.length < maxCols) {
        headers.push(`Column ${headers.length + 1}`);
      }

      // Pad data rows to match maxCols
      const paddedDataRows = dataRows.map((row) => {
        const paddedRow = [...row];
        while (paddedRow.length < maxCols) {
          paddedRow.push("");
        }
        return paddedRow;
      });

      sheetsData[sheetName] = rawJson;
      sheetsMeta[sheetName] = {
        headers,
        rows: paddedDataRows,
        totalRows: paddedDataRows.length,
        totalCols: maxCols,
      };
    });

    // Auto-select active sheet: pick the first sheet that actually has rows
    let primarySheet = workbook.SheetNames[0];
    for (const name of workbook.SheetNames) {
      if (sheetsMeta[name] && sheetsMeta[name].totalRows > 0) {
        primarySheet = name;
        break;
      }
    }

    // Re-order sheetNames so non-empty sheet is first if needed
    const orderedSheetNames = [
      primarySheet,
      ...workbook.SheetNames.filter((s) => s !== primarySheet),
    ];

    return NextResponse.json({
      success: true,
      fileName,
      sizeBytes: file.size,
      sheetNames: orderedSheetNames,
      sheetsData,
      sheetsMeta,
    });
  } catch (error: any) {
    return NextResponse.json(
      { success: false, error: error?.message || "Failed to process Excel spreadsheet" },
      { status: 500 }
    );
  }
}
