"use client";

import { useState, useRef } from "react";
import {
  Upload,
  FileSpreadsheet,
  CheckCircle2,
  AlertCircle,
  RefreshCw,
  Layers,
  FileText,
  Trash2,
  Download,
  Link2,
  Sparkles,
  Search,
  Filter,
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

interface ResultantRecord {
  id: string;
  sheet2Code: string;
  sheet2Name: string;
  sheet1Client: string;
  sheet1Balance: string | number;
  sheet2DueAmount: string | number;
  sheet2Status: string;
  matchType: "Code Match" | "Name Substring Match" | "Token Match" | "Sheet 2 Only (Unmatched)";
  confidence: number;
  isMatched: boolean;
}

interface MatchResultState {
  stats: {
    totalSheet1Rows: number;
    totalSheet2Rows: number;
    resultantTotalRows: number;
    matchedCount: number;
    unmatchedSheet2Count: number;
    unmatchedSheet1Count: number;
    matchRate: number;
  };
  resultantRecords: ResultantRecord[];
  unmatchedSheet1: Array<{ client: string; balance: any }>;
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
  const [selectedView, setSelectedView] = useState<"resultant" | "file1" | "file2">("resultant");
  const [resultFilter, setResultFilter] = useState<"all" | "matched" | "unmatched">("all");

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
        setSelectedView("resultant");
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

  const exportResultantExcel = () => {
    if (!matchResults || matchResults.resultantRecords.length === 0) return;

    let recordsToExport = matchResults.resultantRecords;
    if (resultFilter === "matched") {
      recordsToExport = recordsToExport.filter((r) => r.isMatched);
    } else if (resultFilter === "unmatched") {
      recordsToExport = recordsToExport.filter((r) => !r.isMatched);
    }

    const exportData = recordsToExport.map((m) => ({
      "Customer Code (Sheet 2)": m.sheet2Code,
      "Customer Name (Sheet 2)": m.sheet2Name,
      "Client String (Sheet 1)": m.sheet1Client,
      "Balance (Sheet 1)": m.sheet1Balance,
      "Due Amount (Sheet 2)": m.sheet2DueAmount,
      "Status": m.sheet2Status,
      "Match Status": m.isMatched ? `Matched (${m.matchType})` : "Unmatched (Sheet 2 Only)",
      "Match Confidence": `${m.confidence}%`,
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Resultant All Customers");
    XLSX.writeFile(workbook, `Resultant_All_Sheet2_Customers.xlsx`);
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

  const filteredResultant = matchResults
    ? matchResults.resultantRecords.filter((m) => {
        const matchesQuery =
          m.sheet2Name.toLowerCase().includes(searchQuery.toLowerCase()) ||
          m.sheet1Client.toLowerCase().includes(searchQuery.toLowerCase()) ||
          m.sheet2Code.toLowerCase().includes(searchQuery.toLowerCase());

        if (!matchesQuery) return false;
        if (resultFilter === "matched") return m.isMatched;
        if (resultFilter === "unmatched") return !m.isMatched;
        return true;
      })
    : [];

  return (
    <div className="space-y-8 animate-in fade-in duration-300">
      {/* Title Header */}
      <div className="text-center max-w-2xl mx-auto space-y-2">
        <Badge variant="outline" className="text-emerald-400 border-emerald-500/30 bg-emerald-500/10 mb-2">
          <Link2 className="h-3.5 w-3.5 mr-1" /> Excel Sheet Combiner (Sheet 2 Centric)
        </Badge>
        <h2 className="text-3xl font-extrabold text-white tracking-tight sm:text-4xl">
          Resultant Sheet with ALL Sheet 2 Customers
        </h2>
        <p className="text-sm text-slate-400">
          Generates a complete resultant dataset containing <strong>all customers from Sheet 2</strong>, combined with matched data from Sheet 1.
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
              Sheet 2 (All Customers Target)
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
                  onClick={() => setSelectedView("resultant")}
                  className={`flex items-center gap-2 rounded-lg px-4 py-2 text-xs font-semibold transition-all ${
                    selectedView === "resultant"
                      ? "bg-gradient-to-r from-emerald-600 to-teal-500 text-white shadow-lg shadow-emerald-500/20"
                      : "text-slate-400 hover:bg-slate-900 hover:text-white"
                  }`}
                >
                  <Sparkles className="h-4 w-4 text-emerald-300" />
                  Resultant Sheet (All {matchResults.stats.resultantTotalRows} Sheet 2 Customers)
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
                  Raw Sheet 1 ({file1.fileName})
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
                  Raw Sheet 2 ({file2.fileName})
                </button>
              )}
            </div>

            {matchResults && selectedView === "resultant" && (
              <Button
                variant="gradientSuccess"
                size="sm"
                onClick={exportResultantExcel}
                className="gap-2 text-xs"
              >
                <Download className="h-4 w-4" /> Download Resultant Excel (.xlsx)
              </Button>
            )}
          </div>

          {/* VIEW: Resultant Combined Sheet */}
          {selectedView === "resultant" && matchResults && (
            <div className="space-y-6">
              {/* Match Stats KPI Banner */}
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
                <Card className="p-4 bg-slate-900/90 border-blue-500/30">
                  <p className="text-[11px] font-semibold uppercase text-blue-400">Resultant Total (Sheet 2)</p>
                  <p className="text-2xl font-bold text-white mt-1 font-mono">
                    {matchResults.stats.resultantTotalRows} Customers
                  </p>
                </Card>

                <Card className="p-4 bg-slate-900/90 border-emerald-500/30">
                  <p className="text-[11px] font-semibold uppercase text-emerald-400 flex items-center gap-1">
                    <CheckCircle2 className="h-3.5 w-3.5" /> Matched with Sheet 1
                  </p>
                  <p className="text-2xl font-bold text-emerald-400 mt-1 font-mono">
                    {matchResults.stats.matchedCount}
                  </p>
                </Card>

                <Card className="p-4 bg-slate-900/90 border-amber-500/30">
                  <p className="text-[11px] font-semibold uppercase text-amber-400">Sheet 2 Only (Unmatched)</p>
                  <p className="text-2xl font-bold text-amber-400 mt-1 font-mono">
                    {matchResults.stats.unmatchedSheet2Count}
                  </p>
                </Card>

                <Card className="p-4 bg-slate-900/90 border-purple-500/30">
                  <p className="text-[11px] font-semibold uppercase text-purple-400">Match Coverage Rate</p>
                  <p className="text-2xl font-bold text-purple-300 mt-1 font-mono">
                    {matchResults.stats.matchRate}%
                  </p>
                </Card>
              </div>

              {/* Filters & Search */}
              <div className="flex flex-col sm:flex-row items-center justify-between gap-4">
                <div className="relative flex-1 w-full">
                  <Search className="absolute left-3 top-3 h-4 w-4 text-slate-500" />
                  <input
                    type="text"
                    placeholder="Search resultant sheet by customer code, name, or client string..."
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    className="w-full h-10 rounded-xl border border-slate-800 bg-slate-900/90 pl-9 pr-4 text-xs text-slate-200 placeholder-slate-500 focus:border-emerald-500 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                  />
                </div>

                <div className="flex items-center gap-1.5 bg-slate-900 p-1 rounded-xl border border-slate-800 shrink-0">
                  <button
                    onClick={() => setResultFilter("all")}
                    className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-all ${
                      resultFilter === "all"
                        ? "bg-slate-800 text-white border border-slate-700"
                        : "text-slate-400 hover:text-white"
                    }`}
                  >
                    All Customers ({matchResults.stats.resultantTotalRows})
                  </button>
                  <button
                    onClick={() => setResultFilter("matched")}
                    className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-all ${
                      resultFilter === "matched"
                        ? "bg-emerald-500/20 text-emerald-400 border border-emerald-500/30"
                        : "text-slate-400 hover:text-white"
                    }`}
                  >
                    Matched ({matchResults.stats.matchedCount})
                  </button>
                  <button
                    onClick={() => setResultFilter("unmatched")}
                    className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-all ${
                      resultFilter === "unmatched"
                        ? "bg-amber-500/20 text-amber-400 border border-amber-500/30"
                        : "text-slate-400 hover:text-white"
                    }`}
                  >
                    Unmatched ({matchResults.stats.unmatchedSheet2Count})
                  </button>
                </div>
              </div>

              {/* Resultant Combined Table */}
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
                    <TableHead className="font-bold text-indigo-400">Match Status</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {filteredResultant.length === 0 ? (
                    <TableRow>
                      <TableCell colSpan={8} className="text-center py-8 text-xs text-slate-400">
                        No records match the current search or filter.
                      </TableCell>
                    </TableRow>
                  ) : (
                    filteredResultant.slice(0, 150).map((row, idx) => (
                      <TableRow key={row.id} className={!row.isMatched ? "bg-amber-950/10" : ""}>
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
                          {row.isMatched ? (
                            row.sheet1Client
                          ) : (
                            <span className="text-amber-400/70 italic">Not Found in Sheet 1</span>
                          )}
                        </TableCell>
                        <TableCell className="font-mono text-xs text-right font-bold text-emerald-400">
                          {row.isMatched ? String(row.sheet1Balance) : "-"}
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
                          {row.isMatched ? (
                            <Badge variant="success" className="text-[10px]">
                              {row.matchType} ({row.confidence}%)
                            </Badge>
                          ) : (
                            <Badge variant="warning" className="text-[10px]">
                              Sheet 2 Only
                            </Badge>
                          )}
                        </TableCell>
                      </TableRow>
                    ))
                  )}
                </TableBody>
              </Table>

              {filteredResultant.length > 150 && (
                <p className="text-center text-xs text-slate-500 py-2">
                  Showing first 150 rows of {filteredResultant.length} total resultant rows. Download Excel file for complete list.
                </p>
              )}
            </div>
          )}

          {/* VIEW: Raw File 1 */}
          {selectedView === "file1" && file1 && renderSheetTable(file1)}

          {/* VIEW: Raw File 2 */}
          {selectedView === "file2" && file2 && renderSheetTable(file2)}
        </div>
      )}
    </div>
  );
}
