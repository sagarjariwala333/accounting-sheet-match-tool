"use client";

import { useState, useRef, useEffect } from "react";
import {
  Upload,
  FileSpreadsheet,
  CheckCircle2,
  AlertCircle,
  RefreshCw,
  Layers,
  FileText,
  ArrowRightLeft,
  Trash2,
  Download,
  Link2,
  Sparkles,
  Search,
} from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { Table, TableHeader, TableBody, TableRow, TableHead, TableCell } from "@/components/ui/table";
import { Badge } from "@/components/ui/badge";
import * as XLSX from "xlsx";

interface SheetMeta {
  headers: string[];
  rows: any[][];
  totalRows: number;
  totalCols: number;
}

interface UploadedFileState {
  file: File;
  fileName: string;
  sizeBytes: number;
  sheetNames: string[];
  sheetsData: Record<string, any[][]>;
  sheetsMeta: Record<string, SheetMeta>;
  activeSheet: string;
}

interface MatchedRecord {
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

interface MatchResultState {
  stats: {
    totalSheet1Rows: number;
    totalSheet2Rows: number;
    matchedCount: number;
    unmatchedSheet1Count: number;
    unmatchedSheet2Count: number;
    matchRate: number;
  };
  matchedRecords: MatchedRecord[];
  unmatchedSheet1: Array<{ client: string; balance: any }>;
  unmatchedSheet2: Array<{ code: string; name: string; status: string; dueAmount: any }>;
}

export default function HomePage() {
  const [file1, setFile1] = useState<UploadedFileState | null>(null);
  const [file2, setFile2] = useState<UploadedFileState | null>(null);
  const [loading1, setLoading1] = useState(false);
  const [loading2, setLoading2] = useState(false);
  const [error1, setError1] = useState<string | null>(null);
  const [error2, setError2] = useState<string | null>(null);

  const [matchResults, setMatchResults] = useState<MatchResultState | null>(null);
  const [matchLoading, setMatchLoading] = useState(false);
  const [searchQuery, setSearchQuery] = useState("");
  const [selectedView, setSelectedView] = useState<"combined" | "file1" | "file2" | "unmatched">("combined");

  const input1Ref = useRef<HTMLInputElement>(null);
  const input2Ref = useRef<HTMLInputElement>(null);

  const processUpload = async (targetFile: File, slot: 1 | 2) => {
    const setLoading = slot === 1 ? setLoading1 : setLoading2;
    const setError = slot === 1 ? setError1 : setError2;
    const setFileState = slot === 1 ? setFile1 : setFile2;

    setLoading(true);
    setError(null);

    const formData = new FormData();
    formData.append("file", targetFile);

    try {
      const res = await fetch("/api/upload", {
        method: "POST",
        body: formData,
      });

      const data = await res.json();
      if (!res.ok || !data.success) {
        throw new Error(data.error || `Failed to process Excel File ${slot}`);
      }

      const newState: UploadedFileState = {
        file: targetFile,
        fileName: data.fileName,
        sizeBytes: data.sizeBytes,
        sheetNames: data.sheetNames,
        sheetsData: data.sheetsData,
        sheetsMeta: data.sheetsMeta,
        activeSheet: data.sheetNames[0] || "",
      };

      setFileState(newState);

      // If both files are now uploaded, perform sheet matching!
      const currentFile1 = slot === 1 ? newState : file1;
      const currentFile2 = slot === 2 ? newState : file2;

      if (currentFile1 && currentFile2) {
        runMatching(currentFile1, currentFile2);
      } else {
        setSelectedView(slot === 1 ? "file1" : "file2");
      }
    } catch (err: any) {
      setError(err.message || "Error processing file.");
    } finally {
      setLoading(false);
    }
  };

  const runMatching = async (f1: UploadedFileState, f2: UploadedFileState) => {
    const meta1 = f1.sheetsMeta[f1.activeSheet];
    const meta2 = f2.sheetsMeta[f2.activeSheet];

    if (!meta1 || !meta2) return;

    setMatchLoading(true);
    try {
      const res = await fetch("/api/match", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          sheet1Rows: meta1.rows,
          headers1: meta1.headers,
          sheet2Rows: meta2.rows,
          headers2: meta2.headers,
        }),
      });

      const data = await res.json();
      if (data.success) {
        setMatchResults(data);
        setSelectedView("combined");
      }
    } catch (err) {
      console.error("Matching error:", err);
    } finally {
      setMatchLoading(false);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>, slot: 1 | 2) => {
    const file = e.target.files?.[0];
    if (file) {
      if (!file.name.match(/\.(xlsx|xls|csv)$/i)) {
        const setError = slot === 1 ? setError1 : setError2;
        setError("Invalid file format. Please select an Excel file (.xlsx).");
        return;
      }
      processUpload(file, slot);
    }
  };

  const clearSlot = (slot: 1 | 2) => {
    if (slot === 1) {
      setFile1(null);
      setError1(null);
      setMatchResults(null);
      if (input1Ref.current) input1Ref.current.value = "";
      if (file2) setSelectedView("file2");
    } else {
      setFile2(null);
      setError2(null);
      setMatchResults(null);
      if (input2Ref.current) input2Ref.current.value = "";
      if (file1) setSelectedView("file1");
    }
  };

  const exportCombinedExcel = () => {
    if (!matchResults || matchResults.matchedRecords.length === 0) return;

    const exportData = matchResults.matchedRecords.map((m) => ({
      "Customer Code": m.sheet2Code,
      "Customer Name (Sheet 2)": m.sheet2Name,
      "Client String (Sheet 1)": m.sheet1Client,
      "Balance (Sheet 1)": m.sheet1Balance,
      "Due Amount (Sheet 2)": m.sheet2DueAmount,
      "Status": m.sheet2Status,
      "Match Method": m.matchType,
      "Match Confidence": `${m.confidence}%`,
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Combined Matched Clients");
    XLSX.writeFile(workbook, `Matched_Combined_Clients.xlsx`);
  };

  const exportExcel = (fileState: UploadedFileState) => {
    const meta = fileState.sheetsMeta[fileState.activeSheet];
    if (!meta) return;

    const exportRows = meta.rows.map((row) => {
      const rowObj: Record<string, any> = {};
      meta.headers.forEach((h, idx) => {
        rowObj[h || `Col ${idx + 1}`] = row[idx];
      });
      return rowObj;
    });

    const worksheet = XLSX.utils.json_to_sheet(exportRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, fileState.activeSheet || "Sheet1");
    XLSX.writeFile(workbook, `Export_${fileState.fileName}`);
  };

  // Helper to render sheet data table
  const renderSheetTable = (fileState: UploadedFileState) => {
    const meta = fileState.sheetsMeta[fileState.activeSheet];
    if (!meta) return null;

    const { headers, rows, totalRows, totalCols } = meta;

    return (
      <div className="space-y-6">
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
          <Card className="p-4 bg-slate-900/60 border-slate-800">
            <p className="text-[11px] font-semibold uppercase text-slate-400">Total Rows</p>
            <p className="text-xl font-bold text-white mt-1 font-mono">{totalRows}</p>
          </Card>
          <Card className="p-4 bg-slate-900/60 border-slate-800">
            <p className="text-[11px] font-semibold uppercase text-slate-400">Total Columns</p>
            <p className="text-xl font-bold text-white mt-1 font-mono">{totalCols}</p>
          </Card>
          <Card className="p-4 bg-slate-900/60 border-slate-800">
            <p className="text-[11px] font-semibold uppercase text-slate-400">Active Sheet</p>
            <p className="text-xl font-bold text-emerald-400 mt-1 truncate">{fileState.activeSheet}</p>
          </Card>
          <Card className="p-4 bg-slate-900/60 border-slate-800">
            <p className="text-[11px] font-semibold uppercase text-slate-400">File Size</p>
            <p className="text-xl font-bold text-blue-400 mt-1 font-mono">
              {(fileState.sizeBytes / 1024).toFixed(2)} KB
            </p>
          </Card>
        </div>

        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-800 pb-3">
          <div className="flex items-center gap-2 overflow-x-auto">
            {fileState.sheetNames.length > 1 && (
              <>
                <span className="text-xs font-semibold text-slate-400 mr-2 flex items-center gap-1">
                  <Layers className="h-3.5 w-3.5 text-emerald-400" /> Sheets:
                </span>
                {fileState.sheetNames.map((name) => (
                  <button
                    key={name}
                    onClick={() => {
                      const updated = { ...fileState, activeSheet: name };
                      if (fileState === file1) setFile1(updated);
                      else setFile2(updated);
                    }}
                    className={`rounded-lg px-3 py-1 text-xs font-medium transition-all ${
                      fileState.activeSheet === name
                        ? "bg-emerald-500/20 text-emerald-400 border border-emerald-500/30"
                        : "text-slate-400 hover:bg-slate-800 hover:text-white"
                    }`}
                  >
                    {name}
                  </button>
                ))}
              </>
            )}
          </div>

          <Button variant="outline" size="sm" onClick={() => exportExcel(fileState)} className="gap-2 text-xs">
            <Download className="h-3.5 w-3.5" /> Export Sheet (.xlsx)
          </Button>
        </div>

        {rows.length === 0 ? (
          <Card className="p-8 text-center text-xs text-slate-400">This sheet is empty.</Card>
        ) : (
          <Table>
            <TableHeader>
              <TableRow>
                <TableHead className="w-12 text-center font-mono">#</TableHead>
                {headers.map((col: any, idx: number) => (
                  <TableHead key={idx} className="font-bold text-white">
                    {String(col || `Column ${idx + 1}`)}
                  </TableHead>
                ))}
              </TableRow>
            </TableHeader>
            <TableBody>
              {rows.slice(0, 50).map((row: any[], rowIdx: number) => (
                <TableRow key={rowIdx}>
                  <TableCell className="text-center font-mono text-xs text-slate-500">
                    {rowIdx + 1}
                  </TableCell>
                  {headers.map((_: any, colIdx: number) => (
                    <TableCell key={colIdx} className="text-slate-200 text-xs font-mono">
                      {row[colIdx] !== undefined && row[colIdx] !== null ? String(row[colIdx]) : ""}
                    </TableCell>
                  ))}
                </TableRow>
              ))}
            </TableBody>
          </Table>
        )}
      </div>
    );
  };

  const filteredMatches = matchResults
    ? matchResults.matchedRecords.filter(
        (m) =>
          m.sheet2Name.toLowerCase().includes(searchQuery.toLowerCase()) ||
          m.sheet1Client.toLowerCase().includes(searchQuery.toLowerCase()) ||
          m.sheet2Code.toLowerCase().includes(searchQuery.toLowerCase())
      )
    : [];

  return (
    <div className="space-y-8 animate-in fade-in duration-300">
      {/* Title Header */}
      <div className="text-center max-w-2xl mx-auto space-y-2">
        <Badge variant="outline" className="text-emerald-400 border-emerald-500/30 bg-emerald-500/10 mb-2">
          <Link2 className="h-3.5 w-3.5 mr-1" /> Excel Sheet Matcher & Combiner
        </Badge>
        <h2 className="text-3xl font-extrabold text-white tracking-tight sm:text-4xl">
          Upload & Match Two Excel Sheets (.xlsx)
        </h2>
        <p className="text-sm text-slate-400">
          Upload Sheet 1 (Client & Balance) and Sheet 2 (Code, Name, Status, Due). The system matches Sheet 2 customer names against Sheet 1 client names and combines them.
        </p>
      </div>

      {/* Hidden File Inputs */}
      <input
        ref={input1Ref}
        type="file"
        accept=".xlsx, .xls, .csv"
        onChange={(e) => handleFileChange(e, 1)}
        className="hidden"
      />
      <input
        ref={input2Ref}
        type="file"
        accept=".xlsx, .xls, .csv"
        onChange={(e) => handleFileChange(e, 2)}
        className="hidden"
      />

      {/* Two Upload Slot Cards Side-by-Side */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
        {/* SLOT 1 */}
        <Card className={`relative overflow-hidden p-6 border-2 transition-all duration-300 ${
          file1 ? "border-emerald-500/40 bg-slate-900/90" : "border-dashed border-slate-800 bg-slate-900/40 hover:border-emerald-500/40"
        }`}>
          <div className="flex items-center justify-between mb-4">
            <span className="text-xs font-bold uppercase tracking-wider text-slate-400 flex items-center gap-1.5">
              <span className="flex h-5 w-5 items-center justify-center rounded-full bg-emerald-500/20 text-emerald-400 font-mono text-[11px]">1</span>
              Sheet 1 (Client & Balance)
            </span>
            {file1 && (
              <Badge variant="success" className="gap-1">
                <CheckCircle2 className="h-3 w-3" /> Ready
              </Badge>
            )}
          </div>

          {!file1 ? (
            <div className="text-center py-6 space-y-4">
              <div className="h-12 w-12 mx-auto rounded-xl bg-emerald-500/10 text-emerald-400 flex items-center justify-center border border-emerald-500/20">
                <FileSpreadsheet className="h-6 w-6" />
              </div>
              <div>
                <p className="text-sm font-semibold text-white">Upload Sheet 1</p>
                <p className="text-xs text-slate-400">e.g. EXOTICA HOLIDAYS... - SS28466 | 48,233.00</p>
              </div>

              {error1 && (
                <div className="p-2.5 rounded-lg bg-rose-500/10 text-xs text-rose-400 text-left flex items-center gap-2">
                  <AlertCircle className="h-4 w-4 shrink-0" />
                  <span>{error1}</span>
                </div>
              )}

              <Button
                variant="gradientSuccess"
                disabled={loading1}
                onClick={() => input1Ref.current?.click()}
                className="w-full gap-2 py-5 font-semibold"
              >
                {loading1 ? (
                  <>
                    <RefreshCw className="h-4 w-4 animate-spin" /> Processing Sheet 1...
                  </>
                ) : (
                  <>
                    <Upload className="h-4 w-4" /> Upload Sheet 1 (.xlsx)
                  </>
                )}
              </Button>
            </div>
          ) : (
            <div className="space-y-4">
              <div className="p-3 rounded-xl bg-slate-950/80 border border-slate-800">
                <p className="text-sm font-bold text-white truncate">{file1.fileName}</p>
                <p className="text-xs text-slate-400 mt-0.5">
                  {(file1.sizeBytes / 1024).toFixed(2)} KB &bull; {file1.sheetsMeta[file1.activeSheet]?.totalRows || 0} rows
                </p>
              </div>

              <div className="flex items-center gap-2">
                <Button
                  size="sm"
                  variant={selectedView === "file1" ? "default" : "outline"}
                  onClick={() => setSelectedView("file1")}
                  className="flex-1 text-xs"
                >
                  View Raw Sheet 1
                </Button>
                <Button
                  size="sm"
                  variant="ghost"
                  onClick={() => clearSlot(1)}
                  className="text-slate-400 hover:text-rose-400 hover:bg-rose-500/10"
                >
                  <Trash2 className="h-4 w-4" />
                </Button>
              </div>
            </div>
          )}
        </Card>

        {/* SLOT 2 */}
        <Card className={`relative overflow-hidden p-6 border-2 transition-all duration-300 ${
          file2 ? "border-blue-500/40 bg-slate-900/90" : "border-dashed border-slate-800 bg-slate-900/40 hover:border-blue-500/40"
        }`}>
          <div className="flex items-center justify-between mb-4">
            <span className="text-xs font-bold uppercase tracking-wider text-slate-400 flex items-center gap-1.5">
              <span className="flex h-5 w-5 items-center justify-center rounded-full bg-blue-500/20 text-blue-400 font-mono text-[11px]">2</span>
              Sheet 2 (Code, Name, Status, Due)
            </span>
            {file2 && (
              <Badge variant="default" className="gap-1 bg-blue-500/10 text-blue-400 border-blue-500/20">
                <CheckCircle2 className="h-3 w-3" /> Ready
              </Badge>
            )}
          </div>

          {!file2 ? (
            <div className="text-center py-6 space-y-4">
              <div className="h-12 w-12 mx-auto rounded-xl bg-blue-500/10 text-blue-400 flex items-center justify-center border border-blue-500/20">
                <FileSpreadsheet className="h-6 w-6" />
              </div>
              <div>
                <p className="text-sm font-semibold text-white">Upload Sheet 2</p>
                <p className="text-xs text-slate-400">e.g. SS80497 | Divine tours... | active | 0</p>
              </div>

              {error2 && (
                <div className="p-2.5 rounded-lg bg-rose-500/10 text-xs text-rose-400 text-left flex items-center gap-2">
                  <AlertCircle className="h-4 w-4 shrink-0" />
                  <span>{error2}</span>
                </div>
              )}

              <Button
                variant="default"
                disabled={loading2}
                onClick={() => input2Ref.current?.click()}
                className="w-full gap-2 py-5 font-semibold bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-500 hover:to-indigo-500 shadow-blue-500/20"
              >
                {loading2 ? (
                  <>
                    <RefreshCw className="h-4 w-4 animate-spin" /> Processing Sheet 2...
                  </>
                ) : (
                  <>
                    <Upload className="h-4 w-4" /> Upload Sheet 2 (.xlsx)
                  </>
                )}
              </Button>
            </div>
          ) : (
            <div className="space-y-4">
              <div className="p-3 rounded-xl bg-slate-950/80 border border-slate-800">
                <p className="text-sm font-bold text-white truncate">{file2.fileName}</p>
                <p className="text-xs text-slate-400 mt-0.5">
                  {(file2.sizeBytes / 1024).toFixed(2)} KB &bull; {file2.sheetsMeta[file2.activeSheet]?.totalRows || 0} rows
                </p>
              </div>

              <div className="flex items-center gap-2">
                <Button
                  size="sm"
                  variant={selectedView === "file2" ? "default" : "outline"}
                  onClick={() => setSelectedView("file2")}
                  className="flex-1 text-xs"
                >
                  View Raw Sheet 2
                </Button>
                <Button
                  size="sm"
                  variant="ghost"
                  onClick={() => clearSlot(2)}
                  className="text-slate-400 hover:text-rose-400 hover:bg-rose-500/10"
                >
                  <Trash2 className="h-4 w-4" />
                </Button>
              </div>
            </div>
          )}
        </Card>
      </div>

      {/* Main Switcher Tabs when files are uploaded */}
      {(file1 || file2) && (
        <div className="space-y-6">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-800 pb-3">
            <div className="flex items-center gap-2 overflow-x-auto">
              {matchResults && (
                <button
                  onClick={() => setSelectedView("combined")}
                  className={`flex items-center gap-2 rounded-lg px-4 py-2 text-xs font-semibold transition-all ${
                    selectedView === "combined"
                      ? "bg-gradient-to-r from-emerald-600 to-teal-500 text-white shadow-lg shadow-emerald-500/20"
                      : "text-slate-400 hover:bg-slate-900 hover:text-white"
                  }`}
                >
                  <Sparkles className="h-4 w-4 text-emerald-300" />
                  Combined Matched Sheet ({matchResults.stats.matchedCount})
                </button>
              )}

              {file1 && (
                <button
                  onClick={() => setSelectedView("file1")}
                  className={`flex items-center gap-2 rounded-lg px-4 py-2 text-xs font-semibold transition-all ${
                    selectedView === "file1"
                      ? "bg-slate-800 text-white border border-slate-700"
                      : "text-slate-400 hover:bg-slate-900 hover:text-white"
                  }`}
                >
                  <FileText className="h-4 w-4" />
                  Sheet 1 ({file1.fileName})
                </button>
              )}

              {file2 && (
                <button
                  onClick={() => setSelectedView("file2")}
                  className={`flex items-center gap-2 rounded-lg px-4 py-2 text-xs font-semibold transition-all ${
                    selectedView === "file2"
                      ? "bg-slate-800 text-white border border-slate-700"
                      : "text-slate-400 hover:bg-slate-900 hover:text-white"
                  }`}
                >
                  <FileText className="h-4 w-4" />
                  Sheet 2 ({file2.fileName})
                </button>
              )}

              {matchResults && (matchResults.unmatchedSheet1.length > 0 || matchResults.unmatchedSheet2.length > 0) && (
                <button
                  onClick={() => setSelectedView("unmatched")}
                  className={`flex items-center gap-2 rounded-lg px-4 py-2 text-xs font-semibold transition-all ${
                    selectedView === "unmatched"
                      ? "bg-rose-500/20 text-rose-400 border border-rose-500/30"
                      : "text-slate-400 hover:bg-slate-900 hover:text-white"
                  }`}
                >
                  <AlertCircle className="h-4 w-4 text-rose-400" />
                  Unmatched Items
                </button>
              )}
            </div>

            {matchResults && selectedView === "combined" && (
              <Button
                variant="gradientSuccess"
                size="sm"
                onClick={exportCombinedExcel}
                className="gap-2 text-xs"
              >
                <Download className="h-4 w-4" /> Download Combined Excel (.xlsx)
              </Button>
            )}
          </div>

          {/* VIEW: Combined Matched View */}
          {selectedView === "combined" && matchResults && (
            <div className="space-y-6">
              {/* Match Stats KPI Banner */}
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
                <Card className="p-4 bg-slate-900/90 border-emerald-500/30">
                  <p className="text-[11px] font-semibold uppercase text-emerald-400 flex items-center gap-1">
                    <CheckCircle2 className="h-3.5 w-3.5" /> Matched Records
                  </p>
                  <p className="text-2xl font-bold text-emerald-400 mt-1 font-mono">
                    {matchResults.stats.matchedCount}
                  </p>
                </Card>

                <Card className="p-4 bg-slate-900/90 border-blue-500/30">
                  <p className="text-[11px] font-semibold uppercase text-blue-400">Match Success Rate</p>
                  <p className="text-2xl font-bold text-white mt-1 font-mono">
                    {matchResults.stats.matchRate}%
                  </p>
                </Card>

                <Card className="p-4 bg-slate-900/90 border-rose-500/30">
                  <p className="text-[11px] font-semibold uppercase text-rose-400">Unmatched Sheet 1</p>
                  <p className="text-2xl font-bold text-rose-400 mt-1 font-mono">
                    {matchResults.stats.unmatchedSheet1Count}
                  </p>
                </Card>

                <Card className="p-4 bg-slate-900/90 border-purple-500/30">
                  <p className="text-[11px] font-semibold uppercase text-purple-400">Unmatched Sheet 2</p>
                  <p className="text-2xl font-bold text-purple-300 mt-1 font-mono">
                    {matchResults.stats.unmatchedSheet2Count}
                  </p>
                </Card>
              </div>

              {/* Search Bar */}
              <div className="relative">
                <Search className="absolute left-3 top-3 h-4 w-4 text-slate-500" />
                <input
                  type="text"
                  placeholder="Filter matched records by client name or code..."
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  className="w-full h-10 rounded-xl border border-slate-800 bg-slate-900/90 pl-9 pr-4 text-xs text-slate-200 placeholder-slate-500 focus:border-emerald-500 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                />
              </div>

              {/* Combined Table */}
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead className="w-12 text-center font-mono">#</TableHead>
                    <TableHead className="font-bold text-emerald-400">Code (Sheet 2)</TableHead>
                    <TableHead className="font-bold text-white">Customer Name (Sheet 2)</TableHead>
                    <TableHead className="font-bold text-slate-300">Client String (Sheet 1)</TableHead>
                    <TableHead className="font-bold text-emerald-400 text-right">Balance (Sheet 1)</TableHead>
                    <TableHead className="font-bold text-blue-400 text-right">Due Amount (Sheet 2)</TableHead>
                    <TableHead className="font-bold text-purple-400">Status</TableHead>
                    <TableHead className="font-bold text-indigo-400">Match Method</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {filteredMatches.length === 0 ? (
                    <TableRow>
                      <TableCell colSpan={8} className="text-center py-8 text-xs text-slate-400">
                        No matched records found.
                      </TableCell>
                    </TableRow>
                  ) : (
                    filteredMatches.slice(0, 100).map((row, idx) => (
                      <TableRow key={row.id}>
                        <TableCell className="text-center font-mono text-xs text-slate-500">
                          {idx + 1}
                        </TableCell>
                        <TableCell className="font-mono text-xs text-emerald-400 font-bold">
                          {row.sheet2Code || "N/A"}
                        </TableCell>
                        <TableCell className="font-semibold text-white text-xs">
                          {row.sheet2Name}
                        </TableCell>
                        <TableCell className="font-mono text-xs text-slate-300">
                          {row.sheet1Client}
                        </TableCell>
                        <TableCell className="font-mono text-xs text-right font-bold text-emerald-400">
                          {String(row.sheet1Balance)}
                        </TableCell>
                        <TableCell className="font-mono text-xs text-right font-bold text-blue-400">
                          {String(row.sheet2DueAmount)}
                        </TableCell>
                        <TableCell>
                          <Badge variant="outline" className="text-xs uppercase">
                            {row.sheet2Status && isNaN(Number(row.sheet2Status)) ? row.sheet2Status : "active"}
                          </Badge>
                        </TableCell>
                        <TableCell>
                          <Badge
                            variant={row.matchType === "Code Match" ? "success" : "default"}
                            className="text-[10px]"
                          >
                            {row.matchType} ({row.confidence}%)
                          </Badge>
                        </TableCell>
                      </TableRow>
                    ))
                  )}
                </TableBody>
              </Table>
            </div>
          )}

          {/* VIEW: Raw File 1 */}
          {selectedView === "file1" && file1 && renderSheetTable(file1)}

          {/* VIEW: Raw File 2 */}
          {selectedView === "file2" && file2 && renderSheetTable(file2)}

          {/* VIEW: Unmatched Items */}
          {selectedView === "unmatched" && matchResults && (
            <div className="space-y-6">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                {/* Unmatched Sheet 1 */}
                <Card className="p-5 border-rose-900/40 bg-slate-900/90">
                  <h3 className="text-base font-bold text-rose-400 mb-3 flex items-center gap-2">
                    <AlertCircle className="h-4 w-4" /> Unmatched Sheet 1 ({matchResults.unmatchedSheet1.length})
                  </h3>
                  <div className="max-h-96 overflow-y-auto space-y-2">
                    {matchResults.unmatchedSheet1.map((item, idx) => (
                      <div key={idx} className="p-2.5 rounded-lg bg-slate-950/80 border border-slate-800 text-xs font-mono flex justify-between">
                        <span className="text-white truncate">{item.client}</span>
                        <span className="text-emerald-400 ml-2">{String(item.balance)}</span>
                      </div>
                    ))}
                  </div>
                </Card>

                {/* Unmatched Sheet 2 */}
                <Card className="p-5 border-purple-900/40 bg-slate-900/90">
                  <h3 className="text-base font-bold text-purple-400 mb-3 flex items-center gap-2">
                    <AlertCircle className="h-4 w-4" /> Unmatched Sheet 2 ({matchResults.unmatchedSheet2.length})
                  </h3>
                  <div className="max-h-96 overflow-y-auto space-y-2">
                    {matchResults.unmatchedSheet2.map((item, idx) => (
                      <div key={idx} className="p-2.5 rounded-lg bg-slate-950/80 border border-slate-800 text-xs font-mono flex justify-between">
                        <div>
                          <span className="text-purple-300 font-bold mr-2">[{item.code}]</span>
                          <span className="text-white">{item.name}</span>
                        </div>
                        <span className="text-blue-400 ml-2">{String(item.dueAmount)}</span>
                      </div>
                    ))}
                  </div>
                </Card>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
