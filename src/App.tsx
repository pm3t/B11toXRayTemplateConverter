import React, { useState, useRef } from 'react';
import {
  FileSpreadsheet,
  UploadCloud,
  Download,
  Plane,
  Calendar,
  FileText,
  MapPin,
  Package,
  Search,
  RefreshCw,
  CheckCircle2,
  AlertCircle,
  ShieldCheck,
} from 'lucide-react';
import { parseManifestFile, generateExcelBlob, ConversionResult } from './converter';

export default function App() {
  const [data, setData] = useState<ConversionResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [search, setSearch] = useState('');
  const [isDragging, setIsDragging] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFile = async (file: File) => {
    setError(null);
    setLoading(true);
    try {
      const validExtensions = ['.xlsx', '.xls'];
      const fileExt = '.' + file.name.split('.').pop()?.toLowerCase();
      if (!validExtensions.includes(fileExt)) {
        throw new Error('Format file tidak didukung! Harap unggah file Excel (.xlsx atau .xls).');
      }

      const result = await parseManifestFile(file);
      if (result.rows.length === 0) {
        throw new Error('Tidak ditemukan data pada sheet Detil / Detail di dalam file manifes.');
      }
      setData(result);
    } catch (err: any) {
      console.error(err);
      setError(err?.message || 'Gagal memproses file Excel.');
    } finally {
      setLoading(false);
    }
  };

  const onDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleFile(e.dataTransfer.files[0]);
    }
  };

  const handleDownload = async () => {
    if (!data) return;
    try {
      const blob = await generateExcelBlob(data);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      const baseName = data.fileName.replace(/\.[^/.]+$/, '');
      a.href = url;
      a.download = `OUT_${baseName}.xlsx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err: any) {
      alert('Gagal mendownload file: ' + err.message);
    }
  };

  const filteredRows = data
    ? data.rows.filter((r) => {
        const q = search.toLowerCase();
        return (
          r.mawb.toLowerCase().includes(q) ||
          r.hawb.toLowerCase().includes(q) ||
          r.natureOfGoods.toLowerCase().includes(q) ||
          r.shipper.toLowerCase().includes(q) ||
          r.consignee.toLowerCase().includes(q) ||
          r.origin.toLowerCase().includes(q) ||
          r.destination.toLowerCase().includes(q)
        );
      })
    : [];

  return (
    <div className="min-h-screen bg-slate-900 text-slate-100 flex flex-col font-sans selection:bg-indigo-500 selection:text-white">
      {/* Top Navbar */}
      <header className="border-b border-slate-800 bg-slate-900/80 backdrop-blur sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
          <div className="flex items-center space-x-3">
            <div className="w-10 h-10 rounded-xl bg-gradient-to-tr from-indigo-500 to-cyan-400 flex items-center justify-center shadow-lg shadow-indigo-500/20">
              <FileSpreadsheet className="w-6 h-6 text-white" />
            </div>
            <div>
              <h1 className="font-bold text-lg leading-tight bg-gradient-to-r from-white via-slate-200 to-slate-400 bg-clip-text text-transparent">
                X-Ray Breakdown Report Generator
              </h1>
              <p className="text-xs text-slate-400">Manifes BC 1.1 (.xlsx / .xls) &rarr; OUT.xlsx</p>
            </div>
          </div>
          <div className="flex items-center space-x-2">
            <span className="hidden sm:inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-medium bg-emerald-500/10 text-emerald-400 border border-emerald-500/20">
              <ShieldCheck className="w-3.5 h-3.5" /> Client-Side (100% Aman &amp; Privat)
            </span>
          </div>
        </div>
      </header>

      {/* Main Container */}
      <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-6">
        {/* Upload Box */}
        {!data && (
          <div className="max-w-3xl mx-auto mt-8">
            <div
              onDragOver={(e) => {
                e.preventDefault();
                setIsDragging(true);
              }}
              onDragLeave={() => setIsDragging(false)}
              onDrop={onDrop}
              onClick={() => fileInputRef.current?.click()}
              className={`cursor-pointer border-2 border-dashed rounded-3xl p-12 text-center transition-all duration-200 ${
                isDragging
                  ? 'border-indigo-400 bg-indigo-500/10 scale-[1.01]'
                  : 'border-slate-700 hover:border-indigo-500/60 bg-slate-800/40 hover:bg-slate-800/70'
              }`}
            >
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls"
                className="hidden"
                onChange={(e) => {
                  if (e.target.files && e.target.files[0]) {
                    handleFile(e.target.files[0]);
                  }
                }}
              />
              <div className="w-20 h-20 mx-auto mb-6 rounded-2xl bg-gradient-to-br from-indigo-500/20 to-cyan-500/20 border border-indigo-500/30 flex items-center justify-center text-indigo-400 group-hover:scale-110 transition-transform">
                <UploadCloud className="w-10 h-10" />
              </div>
              <h2 className="text-xl font-semibold mb-2">Tarik &amp; Lepas File Manifes di Sini</h2>
              <p className="text-slate-400 text-sm mb-4">
                Mendukung format <span className="font-semibold text-slate-200">.xlsx</span> dan format Excel lawas{' '}
                <span className="font-semibold text-slate-200">.xls</span>
              </p>
              <div className="inline-flex items-center gap-2 px-5 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-white font-medium text-sm shadow-lg shadow-indigo-600/30 transition-all">
                Pilih File dari Komputer
              </div>
            </div>

            {loading && (
              <div className="mt-6 flex items-center justify-center gap-3 text-indigo-400">
                <RefreshCw className="w-5 h-5 animate-spin" />
                <span className="text-sm font-medium">Sedang memproses dan mengekstrak data manifes...</span>
              </div>
            )}

            {error && (
              <div className="mt-6 p-4 rounded-2xl bg-rose-500/10 border border-rose-500/20 text-rose-300 flex items-start gap-3">
                <AlertCircle className="w-5 h-5 text-rose-400 shrink-0 mt-0.5" />
                <div className="text-sm">{error}</div>
              </div>
            )}

            {/* Feature Highlights */}
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mt-12 text-center sm:text-left">
              <div className="p-5 rounded-2xl bg-slate-800/40 border border-slate-800">
                <div className="w-8 h-8 rounded-lg bg-indigo-500/20 text-indigo-400 flex items-center justify-center mb-3">
                  <CheckCircle2 className="w-5 h-5" />
                </div>
                <h3 className="font-semibold text-sm mb-1">Filter Otomatis</h3>
                <p className="text-xs text-slate-400">
                  Secara cerdas memfilter barang "CONSOL" dan mengambil uraian barang spesifik (dibatasi 50 karakter).
                </p>
              </div>
              <div className="p-5 rounded-2xl bg-slate-800/40 border border-slate-800">
                <div className="w-8 h-8 rounded-lg bg-cyan-500/20 text-cyan-400 flex items-center justify-center mb-3">
                  <Package className="w-5 h-5" />
                </div>
                <h3 className="font-semibold text-sm mb-1">Act Pcs &amp; Tot Pcs</h3>
                <p className="text-xs text-slate-400">
                  Kolom Act Pcs terisi otomatis dari JUMLAH KEMASAN dan Tot Pcs dari TOTAL KEMASAN.
                </p>
              </div>
              <div className="p-5 rounded-2xl bg-slate-800/40 border border-slate-800">
                <div className="w-8 h-8 rounded-lg bg-emerald-500/20 text-emerald-400 flex items-center justify-center mb-3">
                  <Download className="w-5 h-5" />
                </div>
                <h3 className="font-semibold text-sm mb-1">Siap untuk Vercel</h3>
                <p className="text-xs text-slate-400">
                  Berjalan penuh di sisi browser tanpa batasan ukuran file, tanpa latency, dan tanpa server backend.
                </p>
              </div>
            </div>
          </div>
        )}

        {/* Dashboard View (when data is loaded) */}
        {data && (
          <div className="space-y-6">
            {/* Action Bar */}
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 bg-slate-800/70 p-4 rounded-2xl border border-slate-700/60">
              <div>
                <span className="text-xs text-slate-400 uppercase tracking-wider font-semibold">File Sumber</span>
                <div className="flex items-center gap-2 mt-0.5">
                  <h2 className="text-base font-bold text-white flex items-center gap-2">
                    <FileSpreadsheet className="w-4 h-4 text-indigo-400" />
                    {data.fileName}
                  </h2>
                  <span className={`text-[11px] px-2.5 py-0.5 rounded-full font-semibold border ${
                    data.formatType === 'NVOCC' 
                      ? 'bg-purple-500/10 text-purple-300 border-purple-500/20' 
                      : 'bg-blue-500/10 text-blue-300 border-blue-500/20'
                  }`}>
                    {data.formatType === 'NVOCC' ? 'Format: NVOCC / BL' : 'Format: Manifes BC 1.1'}
                  </span>
                </div>
              </div>
              <div className="flex items-center gap-3">
                <button
                  onClick={() => {
                    setData(null);
                    setSearch('');
                  }}
                  className="px-4 py-2 rounded-xl bg-slate-700 hover:bg-slate-600 text-slate-200 text-sm font-medium flex items-center gap-2 transition"
                >
                  <RefreshCw className="w-4 h-4" /> Ganti File
                </button>
                <button
                  onClick={handleDownload}
                  className="px-5 py-2 rounded-xl bg-gradient-to-r from-indigo-500 to-cyan-500 hover:from-indigo-400 hover:to-cyan-400 text-white text-sm font-semibold shadow-lg shadow-indigo-500/25 flex items-center gap-2 transition"
                >
                  <Download className="w-4 h-4" /> Download OUT.xlsx
                </button>
              </div>
            </div>

            {/* Flight / Header Cards */}
            <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-3">
              <div className="p-4 rounded-2xl bg-slate-800/50 border border-slate-700/60">
                <div className="flex items-center gap-2 text-slate-400 text-xs mb-1">
                  <Plane className="w-3.5 h-3.5 text-indigo-400" /> Penerbangan
                </div>
                <div className="text-base font-bold text-white">{data.header.flightNumber || '-'}</div>
              </div>

              <div className="p-4 rounded-2xl bg-slate-800/50 border border-slate-700/60">
                <div className="flex items-center gap-2 text-slate-400 text-xs mb-1">
                  <Calendar className="w-3.5 h-3.5 text-cyan-400" /> Tgl Kedatangan
                </div>
                <div className="text-base font-bold text-white">
                  {data.header.flightDate instanceof Date
                    ? data.header.flightDate.toLocaleDateString('id-ID')
                    : '-'}
                </div>
              </div>

              <div className="p-4 rounded-2xl bg-slate-800/50 border border-slate-700/60">
                <div className="flex items-center gap-2 text-slate-400 text-xs mb-1">
                  <FileText className="w-3.5 h-3.5 text-amber-400" /> Nomor BC 1.1
                </div>
                <div className="text-base font-bold text-white">{data.header.noBc11 || '-'}</div>
              </div>

              <div className="p-4 rounded-2xl bg-slate-800/50 border border-slate-700/60">
                <div className="flex items-center gap-2 text-slate-400 text-xs mb-1">
                  <MapPin className="w-3.5 h-3.5 text-rose-400" /> Rute
                </div>
                <div className="text-base font-bold text-white">
                  {data.header.origin || '-'} &rarr; {data.header.destination || '-'}
                </div>
              </div>

              <div className="p-4 rounded-2xl bg-slate-800/50 border border-slate-700/60 col-span-2 sm:col-span-1">
                <div className="flex items-center gap-2 text-slate-400 text-xs mb-1">
                  <Package className="w-3.5 h-3.5 text-emerald-400" /> Total Pos Screening
                </div>
                <div className="text-base font-bold text-white flex items-baseline gap-2">
                  <span>{data.totalRecords} Pos</span>
                  {data.transshipmentCount > 0 && (
                    <span className="text-xs font-normal text-amber-400">
                      ({data.transshipmentCount} transshipment dilewati)
                    </span>
                  )}
                </div>
              </div>
            </div>

            {/* Search Filter & Table Container */}
            <div className="bg-slate-800/50 rounded-2xl border border-slate-700/60 overflow-hidden">
              <div className="p-4 border-b border-slate-700/60 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                <div className="relative flex-1 max-w-md">
                  <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                  <input
                    type="text"
                    value={search}
                    onChange={(e) => setSearch(e.target.value)}
                    placeholder="Cari berdasarkan MAWB, HAWB, Shipper, atau Barang..."
                    className="w-full pl-9 pr-4 py-2 bg-slate-900 border border-slate-700 rounded-xl text-sm text-slate-200 placeholder-slate-500 focus:outline-none focus:border-indigo-500"
                  />
                </div>
                <div className="text-xs text-slate-400">
                  Menampilkan <span className="font-bold text-slate-200">{filteredRows.length}</span> dari{' '}
                  <span className="font-bold text-slate-200">{data.totalRecords}</span> baris
                </div>
              </div>

              {/* Table */}
              <div className="overflow-x-auto max-h-[500px]">
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="bg-slate-900/90 text-slate-400 uppercase tracking-wider text-[11px] sticky top-0 z-10 border-b border-slate-700">
                    <tr>
                      <th className="py-3 px-3">No</th>
                      <th className="py-3 px-3">MAWB</th>
                      <th className="py-3 px-3">HAWB</th>
                      <th className="py-3 px-3 text-center">Act Pcs</th>
                      <th className="py-3 px-3 text-center">Tot Pcs</th>
                      <th className="py-3 px-3 text-right">Act Weight</th>
                      <th className="py-3 px-3 text-right">Tot Weight</th>
                      <th className="py-3 px-3">Nature of Goods (Max 50)</th>
                      <th className="py-3 px-3">Shipper</th>
                      <th className="py-3 px-3">Consignee</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-700/50">
                    {filteredRows.map((r) => (
                      <tr key={r.no} className="hover:bg-slate-700/30 transition">
                        <td className="py-2.5 px-3 font-semibold text-slate-400">{r.no}</td>
                        <td className="py-2.5 px-3 font-mono font-medium text-indigo-300">{r.mawb}</td>
                        <td className="py-2.5 px-3 font-mono text-slate-300">{r.hawb}</td>
                        <td className="py-2.5 px-3 text-center font-bold text-cyan-400 bg-cyan-500/5">
                          {r.actPcs ?? '-'}
                        </td>
                        <td className="py-2.5 px-3 text-center font-bold text-emerald-400 bg-emerald-500/5">
                          {r.totPcs ?? '-'}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono text-slate-300">
                          {r.actWeight || '-'}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono text-slate-300">{r.totWeight}</td>
                        <td className="py-2.5 px-3 text-amber-200 max-w-xs truncate" title={r.natureOfGoods}>
                          {r.natureOfGoods || '-'}
                        </td>
                        <td className="py-2.5 px-3 max-w-[200px] truncate text-slate-300" title={r.shipper}>
                          {r.shipper}
                        </td>
                        <td className="py-2.5 px-3 max-w-[200px] truncate text-slate-300" title={r.consignee}>
                          {r.consignee}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}
      </main>

      {/* Footer */}
      <footer className="border-t border-slate-800 py-6 text-center text-xs text-slate-500">
        X-Ray Template Generator &bull; Siap dideploy di Vercel &bull; Bebas server &amp; berjalan langsung di browser
      </footer>
    </div>
  );
}

