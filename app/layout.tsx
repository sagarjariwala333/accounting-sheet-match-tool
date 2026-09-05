import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Excel (.xlsx) Sheet Upload & Viewer",
  description: "Simple Next.js application to upload and inspect Excel sheets (.xlsx).",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en" className="dark">
      <body className="antialiased min-h-screen bg-slate-950 text-slate-100 selection:bg-emerald-500 selection:text-white">
        <div className="min-h-screen flex flex-col">
          <header className="border-b border-slate-800 bg-slate-900/60 backdrop-blur-md px-6 py-4 flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="h-9 w-9 rounded-xl bg-gradient-to-tr from-emerald-600 to-teal-400 flex items-center justify-center font-black text-white text-lg shadow-md shadow-emerald-500/20">
                X
              </div>
              <div>
                <h1 className="text-base font-bold text-white tracking-tight">Excel Tool</h1>
                <p className="text-xs text-slate-400">XLSX File Parser & Uploader</p>
              </div>
            </div>
            <span className="text-xs font-mono rounded-full bg-slate-800 px-3 py-1 text-slate-400 border border-slate-700">
              Next.js Full-Stack API
            </span>
          </header>
          <main className="flex-1 p-6 md:p-12 max-w-6xl mx-auto w-full">
            {children}
          </main>
        </div>
      </body>
    </html>
  );
}
