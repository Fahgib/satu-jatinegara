"use client";

import React, { useState, useEffect, useRef, useMemo } from "react";
import {
  Upload,
  FileSpreadsheet,
  Presentation,
  PlusCircle,
  RotateCcw,
  TrendingUp,
  TrendingDown,
  ArrowRight,
  AlertCircle,
  CheckCircle2,
  Clock,
  ShieldCheck,
  Filter,
  Flame,
  Search,
  Printer,
  CheckSquare,
  Square,
  FileText,
  Building2,
} from "lucide-react";
import { MetricCard } from "@/components/ui/MetricCard";
import { parseExcelKogolBuffer, HasilKogol } from "@/lib/parserKogol";
import { downloadPresentationPptx } from "@/lib/exportPptx";
import * as XLSX from "xlsx";
import {
  BarChart,
  Bar,
  PieChart,
  Pie,
  Cell,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
} from "recharts";

const PlnLogo = ({ className = "w-9 h-11" }: { className?: string }) => (
  <svg viewBox="0 0 100 125" className={className} fill="none" xmlns="http://www.w3.org/2000/svg">
    <path d="M10 8 H90 L80 115 L20 115 Z" fill="#FFDD00" stroke="#F59E0B" strokeWidth="2" />
    <path d="M24 78 C32 74, 40 82, 48 78 C56 74, 64 82, 74 78" stroke="#005C8A" strokeWidth="4.5" strokeLinecap="round" fill="none" />
    <path d="M23 88 C31 84, 39 92, 47 88 C55 84, 63 92, 75 88" stroke="#005C8A" strokeWidth="4.5" strokeLinecap="round" fill="none" />
    <path d="M22 98 C30 94, 38 102, 46 98 C54 94, 62 102, 76 98" stroke="#005C8A" strokeWidth="4.5" strokeLinecap="round" fill="none" />
    <path d="M58 14 L32 54 L52 54 L42 82 L72 42 L52 42 Z" fill="#E11D48" stroke="#BE123C" strokeWidth="1.5" strokeLinejoin="round" />
  </svg>
);

type SkenarioSimulasi = "SEMUA" | "TANPA_KCIC" | "HANYA_NON_AMR";

export default function DashboardPage() {
  const [dataHistory, setDataHistory] = useState<Record<string, HasilKogol>>({});
  const [selectedKey, setSelectedKey] = useState<string>("");
  const [modeLaporan] = useState<"BULANAN" | "TAHUNAN">("BULANAN");
  const [waktuSinkronisasi, setWaktuSinkronisasi] = useState<string>("");
  const [skenario, setSkenario] = useState<SkenarioSimulasi>("SEMUA");
  const [searchQuery, setSearchQuery] = useState<string>("");
  const [filterKategori, setFilterKategori] = useState<string>("SEMUA");
  const [simulasiLunas, setSimulasiLunas] = useState<Record<string, boolean>>({});
  const [activeTab, setActiveTab] = useState<"RINGKASAN" | "LEMBAR_KERJA">("RINGKASAN");

  const fileInputRef = useRef<HTMLInputElement>(null);
  const targetSaldo = 982000000;

  useEffect(() => {
    const now = new Date();
    const formatted = new Intl.DateTimeFormat("id-ID", {
      day: "2-digit",
      month: "long",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit",
      hour12: false,
    }).format(now);
    setWaktuSinkronisasi(`${formatted.replace(" pukul", ",")} WIB`);

    const cached = localStorage.getItem("satu_jatinegara_history");
    if (cached) {
      try {
        const parsed = JSON.parse(cached);
        if (Object.keys(parsed).length > 0) {
          setDataHistory(parsed);
          const keys = Object.keys(parsed).sort();
          setSelectedKey(keys[keys.length - 1]);
        }
      } catch (e) {
        console.error("Gagal membaca cache lokal", e);
      }
    }
  }, []);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const buffer = evt.target?.result as ArrayBuffer;
        const parsed = parseExcelKogolBuffer(buffer, file.name);
        const key = `${parsed.tahun}-${parsed.kodeBulan}`;

        setDataHistory((prev) => {
          const updated = { ...prev, [key]: parsed };
          try {
            localStorage.setItem("satu_jatinegara_history", JSON.stringify(updated));
          } catch (storageErr) {
            console.warn("Storage browser penuh:", storageErr);
          }
          return updated;
        });

        setSelectedKey(key);
        setSimulasiLunas({});

        const now = new Date();
        const formatted = new Intl.DateTimeFormat("id-ID", {
          day: "2-digit",
          month: "long",
          year: "numeric",
          hour: "2-digit",
          minute: "2-digit",
          hour12: false,
        }).format(now);
        setWaktuSinkronisasi(`${formatted.replace(" pukul", ",")} WIB`);
      } catch (err) {
        console.error("Gagal memproses file:", err);
        alert("Gagal membaca struktur Excel. Pastikan berkas adalah laporan resmi KOGOL.");
      } finally {
        if (fileInputRef.current) fileInputRef.current.value = "";
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleResetData = () => {
    localStorage.removeItem("satu_jatinegara_history");
    setDataHistory({});
    setSelectedKey("");
    setSimulasiLunas({});
    if (fileInputRef.current) fileInputRef.current.value = "";
  };

  const activeData: HasilKogol | null = dataHistory[selectedKey] || null;
  const sortedKeys = Object.keys(dataHistory).sort();

  const totalLunasSimulasi = useMemo(() => {
    if (!activeData || !activeData.semuaPelanggan) return 0;
    return activeData.semuaPelanggan.reduce((acc, p) => {
      return simulasiLunas[p.idpel] ? acc + p.rpTunggakan : acc;
    }, 0);
  }, [activeData, simulasiLunas]);

  const jumlahPelangganLunasSimulasi = Object.values(simulasiLunas).filter(Boolean).length;

  let saldoSimulasi = activeData ? activeData.denganKcic?.totalRp || 0 : 0;
  let plgnSimulasi = activeData ? activeData.denganKcic?.totalPlgn || 0 : 0;

  if (activeData) {
    if (skenario === "TANPA_KCIC") {
      saldoSimulasi = activeData.tanpaKcic?.totalRp || 0;
      plgnSimulasi = activeData.tanpaKcic?.totalPlgn || 0;
    } else if (skenario === "HANYA_NON_AMR") {
      saldoSimulasi = activeData.denganKcic?.nonAmrRp || 0;
      plgnSimulasi = activeData.denganKcic?.nonAmrPlgn || 0;
    }
  }

  saldoSimulasi = Math.max(0, saldoSimulasi - totalLunasSimulasi);
  plgnSimulasi = Math.max(0, plgnSimulasi - jumlahPelangganLunasSimulasi);

  const capaianPersenSimulasi = (((targetSaldo - saldoSimulasi) / targetSaldo) * 100).toFixed(1);

  const hasComparison = sortedKeys.length >= 2;
  const prevKey = hasComparison ? sortedKeys[sortedKeys.length - 2] : null;
  const currKey = hasComparison ? sortedKeys[sortedKeys.length - 1] : null;

  const prevData = prevKey ? dataHistory[prevKey] : null;
  const currData = currKey ? dataHistory[currKey] : null;

  const diffTotalRp = currData && prevData ? (currData.denganKcic?.totalRp || 0) - (prevData.denganKcic?.totalRp || 0) : 0;
  const pctTotalRp = prevData && (prevData.denganKcic?.totalRp || 0) > 0 ? (diffTotalRp / prevData.denganKcic.totalRp) * 100 : 0;
  const diffTanpaKcicRp = currData && prevData ? (currData.tanpaKcic?.totalRp || 0) - (prevData.tanpaKcic?.totalRp || 0) : 0;
  const pctTanpaKcicRp = prevData && (prevData.tanpaKcic?.totalRp || 0) > 0 ? (diffTanpaKcicRp / prevData.tanpaKcic.totalRp) * 100 : 0;
  const diffPlgn = currData && prevData ? (currData.denganKcic?.totalPlgn || 0) - (prevData.denganKcic?.totalPlgn || 0) : 0;
  const diffKcicRp = currData && prevData ? (currData.kcicOnly?.rp || 0) - (prevData.kcicOnly?.rp || 0) : 0;

  const kcicContributionPct = diffTotalRp > 0 && diffKcicRp > 0
    ? ((diffKcicRp / diffTotalRp) * 100).toFixed(1)
    : "0.0";

  const filteredPelanggan = useMemo(() => {
    if (!activeData || !activeData.semuaPelanggan) return [];
    return activeData.semuaPelanggan.filter((p) => {
      const matchText =
        p.idpel.toLowerCase().includes(searchQuery.toLowerCase()) ||
        p.nama.toLowerCase().includes(searchQuery.toLowerCase());
      const matchKategori = filterKategori === "SEMUA" || p.kategori === filterKategori;
      return matchText && matchKategori;
    });
  }, [activeData, searchQuery, filterKategori]);

  const togglePelunasan = (idpel: string) => {
    setSimulasiLunas((prev) => ({
      ...prev,
      [idpel]: !prev[idpel],
    }));
  };

  const handlePrint = () => {
    window.print();
  };

  const barChartData = sortedKeys
    .map((k) => {
      const item = dataHistory[k];
      if (!item) return null;
      const namaBulanSingkat = (item.periode || "").split(" ")[0]?.substring(0, 3) || k;
      return {
        bulan: `${namaBulanSingkat} '${(item.tahun || "26").substring(2)}`,
        Dengan_KCIC: item.denganKcic?.totalRp || 0,
        Tanpa_KCIC: item.tanpaKcic?.totalRp || 0,
        Target: targetSaldo,
      };
    })
    .filter(Boolean);

  const handleDownloadExcel = () => {
    if (!activeData || sortedKeys.length === 0) return;

    const wb = XLSX.utils.book_new();

    // SHEET 1: RINGKASAN EKSEKUTIF & KOMPARASI
    const sheet1Data: any[][] = [
      ["PT PLN (PERSERO) UID JAKARTA RAYA"],
      ["UP3 JATINEGARA — BIDANG TRANSAKSI ENERGI"],
      ["LAPORAN MONITORING SALDO AKHIR TUNGGAKAN (KOGOL 0)"],
      [`PERIODE CUT-OFF: ${activeData.periode.toUpperCase()}`],
      [`WAKTU GENERATE: ${waktuSinkronisasi}`],
      [],
      ["PARAMETER KINERJA", "NILAI"],
      ["Target Saldo Unit", targetSaldo],
      ["Realisasi Saldo Unit (Dengan KCIC)", activeData.denganKcic.totalRp],
      ["Realisasi Saldo Murni (Tanpa KCIC)", activeData.tanpaKcic.totalRp],
      ["Khusus Entitas KCIC", activeData.kcicOnly.rp],
      ["Pencapaian Terhadap Target", `${capaianPersenSimulasi}%`],
      [],
      ["REKAPITULASI REALISASI SALDO MULTI-PERIODE (RUPIAH)"],
      [
        "Kategori",
        "Jenis Pelanggan",
        ...sortedKeys.map((k) => dataHistory[k].periode),
      ],
      [
        "Dengan KCIC",
        "AMR",
        ...sortedKeys.map((k) => dataHistory[k].denganKcic.amrRp),
      ],
      [
        "Dengan KCIC",
        "NON-AMR",
        ...sortedKeys.map((k) => dataHistory[k].denganKcic.nonAmrRp),
      ],
      [
        "Tanpa KCIC",
        "AMR",
        ...sortedKeys.map((k) => dataHistory[k].tanpaKcic.amrRp),
      ],
      [
        "Tanpa KCIC",
        "NON-AMR",
        ...sortedKeys.map((k) => dataHistory[k].tanpaKcic.nonAmrRp),
      ],
      [
        "Khusus KCIC",
        "AMR KCIC",
        ...sortedKeys.map((k) => dataHistory[k].kcicOnly.rp),
      ],
      [],
      ["SEGMENTASI TUNGGAKAN BERDASARKAN GOLONGAN TARIF"],
      ["Golongan Tarif", "Jumlah Pelanggan", "Nominal Tunggakan (Rp)"],
      ["Rumah Tangga (R)", activeData.rekapTarif?.R?.plgn || 0, activeData.rekapTarif?.R?.rp || 0],
      ["Bisnis (B)", activeData.rekapTarif?.B?.plgn || 0, activeData.rekapTarif?.B?.rp || 0],
      ["Industri (I)", activeData.rekapTarif?.I?.plgn || 0, activeData.rekapTarif?.I?.rp || 0],
      ["Publik / Pemerintahan (P)", activeData.rekapTarif?.P?.plgn || 0, activeData.rekapTarif?.P?.rp || 0],
      ["Traksi Kereta / KCIC (T)", activeData.rekapTarif?.T?.plgn || 0, activeData.rekapTarif?.T?.rp || 0],
      [],
      ["LEMBAR PENGESAHAN LAPORAN"],
      ["Disusun Oleh:", "", "Diperiksa Oleh:", "", "Disetujui Oleh:"],
      ["TL Pengelolaan Piutang", "", "Asman Transaksi Energi", "", "Manager UP3 Jatinegara"],
      [],
      [],
      ["( .................................... )", "", "( .................................... )", "", "( .................................... )"],
    ];

    const ws1 = XLSX.utils.aoa_to_sheet(sheet1Data);
    ws1["!cols"] = [
      { wch: 34 },
      { wch: 22 },
      { wch: 24 },
      { wch: 24 },
      { wch: 24 },
    ];
    XLSX.utils.book_append_sheet(wb, ws1, "Ringkasan Eksekutif");

    // SHEET 2: TOP 10 PARETO PRIORITAS PENAGIHAN
    const top10Data: any[][] = [
      ["TOP 10 PELANGGAN SALDO TUNGGAKAN TERBESAR (PARETO 80/20)"],
      [`PERIODE: ${activeData.periode.toUpperCase()}`],
      [],
      ["No", "ID Pelanggan", "Nama Pelanggan", "Tarif / Daya", "Segmen", "Nominal Tunggakan (Rp)", "Kontribusi (%)"],
    ];

    const top10List = (activeData.semuaPelanggan || []).slice(0, 10);
    top10List.forEach((p, idx) => {
      const kontribusi = ((p.rpTunggakan / (activeData.denganKcic.totalRp || 1)) * 100).toFixed(2);
      top10Data.push([
        idx + 1,
        p.idpel,
        p.nama,
        p.tarifDaya,
        p.kategori,
        p.rpTunggakan,
        `${kontribusi}%`,
      ]);
    });

    const ws2 = XLSX.utils.aoa_to_sheet(top10Data);
    ws2["!cols"] = [
      { wch: 6 },
      { wch: 18 },
      { wch: 35 },
      { wch: 16 },
      { wch: 12 },
      { wch: 24 },
      { wch: 15 },
    ];
    XLSX.utils.book_append_sheet(wb, ws2, "Top 10 Prioritas");

    // SHEET 3: DATA LENGKAP PELANGGAN & LEMBAR KERJA
    const semuaData: any[][] = [
      ["DAFTAR LENGKAP TUNGGAKAN PELANGGAN KOGOL UNIT JATINEGARA"],
      [`PERIODE: ${activeData.periode.toUpperCase()}`],
      [],
      ["No", "ID Pelanggan", "Nama Pelanggan", "Tarif / Daya", "Segmen", "Golongan", "Nominal Tunggakan (Rp)", "Status Lapangan", "Catatan Petugas"],
    ];

    (activeData.semuaPelanggan || []).forEach((p, idx) => {
      semuaData.push([
        idx + 1,
        p.idpel,
        p.nama,
        p.tarifDaya,
        p.kategori,
        p.kelompokTarif,
        p.rpTunggakan,
        "",
        "",
      ]);
    });

    const ws3 = XLSX.utils.aoa_to_sheet(semuaData);
    ws3["!cols"] = [
      { wch: 6 },
      { wch: 18 },
      { wch: 35 },
      { wch: 16 },
      { wch: 12 },
      { wch: 10 },
      { wch: 24 },
      { wch: 18 },
      { wch: 25 },
    ];
    XLSX.utils.book_append_sheet(wb, ws3, "Data Lengkap Pelanggan");

    XLSX.writeFile(
      wb,
      `PLN_UP3_Jatinegara_Laporan_KOGOL_${activeData.periode.replace(/\s+/g, "_")}.xlsx`
    );
  };

  const donutPelanggan = activeData
    ? [
        { name: "AMR", value: activeData.denganKcic?.amrPlgn || 0, color: "#005C8A" },
        { name: "NON-AMR", value: activeData.denganKcic?.nonAmrPlgn || 0, color: "#FFD100" },
      ]
    : [];

  const donutRupiah = activeData
    ? [
        { name: "AMR", value: activeData.denganKcic?.amrRp || 0, color: "#005C8A" },
        { name: "NON-AMR", value: activeData.denganKcic?.nonAmrRp || 0, color: "#EA580C" },
      ]
    : [];

  return (
    <main className="min-h-screen bg-[#F4F7F9] p-4 lg:p-8 space-y-6 print:bg-white print:p-2">
      {/* ============================================================== */}
      {/* KOP RESMI DOKUMEN CETAK PLN (Hanya Tampil Saat Dicetak ke PDF) */}
      {/* ============================================================== */}
      {activeData && (
        <div className="hidden print:block border-b-2 border-slate-900 pb-3 mb-4">
          <div className="flex justify-between items-start">
            <div className="flex items-center gap-3">
              <PlnLogo className="w-12 h-14" />
              <div>
                <h2 className="text-sm font-black text-slate-900 tracking-wider">PT PLN (PERSERO) UID JAKARTA RAYA</h2>
                <h3 className="text-xs font-bold text-slate-800">UP3 JATINEGARA — BIDANG TRANSAKSI ENERGI</h3>
                <p className="text-[10px] text-slate-500">Jl. Jatinegara Timur No. 89, Jakarta Timur | Telp: (021) 819-0021</p>
              </div>
            </div>
            <div className="text-right text-[10px] text-slate-600 space-y-0.5">
              <div>No. Dokumen : <strong>PLN-UP3JTN/KOGOL/{activeData.tahun}/{activeData.kodeBulan}</strong></div>
              <div>Klasifikasi  : <strong>Rahasia / Laporan Manajemen</strong></div>
              <div>Tanggal Cetak: <strong>{waktuSinkronisasi}</strong></div>
            </div>
          </div>
          <div className="text-center mt-3 pt-2 border-t border-slate-200">
            <h1 className="text-sm font-black text-slate-900 uppercase">
              LEMBAR LAPORAN REALISASI SALDO AKHIR TUNGGAKAN (KOGOL 0)
            </h1>
            <p className="text-xs font-semibold text-slate-700">Periode Cut-Off: {activeData.periode.toUpperCase()}</p>
          </div>
        </div>
      )}

      {/* Header Korporat PLN Layar Web */}
      <header className="bg-white rounded-2xl p-6 border border-slate-200 shadow-xs space-y-4 print:hidden">
        <div className="flex flex-col xl:flex-row justify-between xl:items-center gap-5 border-b border-slate-100 pb-5">
          <div className="flex items-center gap-4">
            <div className="p-2 bg-slate-50 border border-slate-200/80 rounded-xl shadow-xs">
              <PlnLogo className="w-10 h-12" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="text-xs font-black tracking-widest text-[#005C8A] uppercase">
                  PT PLN (Persero)
                </span>
                <span className="text-slate-300">•</span>
                <span className="text-xs font-bold text-slate-500">
                  UID Jakarta Raya — UP3 Jatinegara
                </span>
              </div>
              <h1 className="text-2xl font-black text-slate-900 tracking-tight flex items-center gap-2.5 mt-0.5">
                SATU JATINEGARA
                <span className="text-[11px] font-bold px-2 py-0.5 rounded-md bg-sky-50 text-[#005C8A] border border-sky-200">
                  Enterprise Suite
                </span>
              </h1>
              <p className="text-xs text-slate-500 mt-0.5">
                Sistem Monitoring & Otomasi Analisis Saldo Akhir Tunggakan Rekening Listrik (KOGOL)
              </p>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2.5">
            {activeData && (
              <>
                <button
                  onClick={handlePrint}
                  className="flex items-center gap-2 bg-sky-600 hover:bg-sky-700 text-white text-xs font-semibold px-3.5 py-2.5 rounded-lg shadow-xs transition"
                  title="Cetak format resume / dokumen tanda tangan resmi"
                >
                  <Printer className="w-4 h-4" />
                  <span>Cetak Dokumen (PDF)</span>
                </button>

                <button
                  onClick={handleDownloadExcel}
                  className="flex items-center gap-2 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold px-3.5 py-2.5 rounded-lg shadow-xs transition"
                >
                  <FileSpreadsheet className="w-4 h-4" />
                  <span>Excel ({sortedKeys.length} Bln)</span>
                </button>

                <button
                  onClick={() => downloadPresentationPptx(activeData, modeLaporan, dataHistory)}
                  className="flex items-center gap-2 bg-slate-800 hover:bg-slate-900 text-white text-xs font-semibold px-3.5 py-2.5 rounded-lg shadow-xs transition"
                >
                  <Presentation className="w-4 h-4" />
                  <span>Unduh PPTX</span>
                </button>
              </>
            )}

            <label className="cursor-pointer flex items-center gap-2 bg-[#005C8A] hover:bg-[#00476B] text-white text-xs font-semibold px-4 py-2.5 rounded-lg shadow-xs transition">
              {sortedKeys.length === 0 ? (
                <>
                  <Upload className="w-4 h-4" />
                  <span>Unggah Excel KOGOL</span>
                </>
              ) : (
                <>
                  <PlusCircle className="w-4 h-4" />
                  <span>Tambah Bulan</span>
                </>
              )}
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls"
                className="hidden"
                onChange={handleFileUpload}
              />
            </label>

            {sortedKeys.length > 0 && (
              <button
                onClick={handleResetData}
                title="Reset Semua Data"
                className="p-2.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg border border-slate-200 transition"
              >
                <RotateCcw className="w-4 h-4" />
              </button>
            )}
          </div>
        </div>

        <div className="flex flex-col sm:flex-row justify-between sm:items-center gap-3 text-xs">
          <div className="flex flex-wrap items-center gap-2">
            <span className="font-semibold text-slate-500">Status Operasional:</span>
            <span className="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full bg-emerald-50 text-emerald-700 font-bold border border-emerald-200">
              <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span>
              KOGOL 0 AKHIR
            </span>
            <span className="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full bg-sky-50 text-[#005C8A] font-bold border border-sky-200">
              <ShieldCheck className="w-3.5 h-3.5" />
              SIAP AUDIT
            </span>
            {activeData && (
              <span className="px-2.5 py-0.5 rounded-full bg-amber-50 text-amber-800 font-semibold border border-amber-200">
                Fokus Periode: <b>{activeData.periode}</b>
              </span>
            )}
          </div>

          <div className="flex items-center gap-1.5 text-slate-500 font-medium">
            <Clock className="w-3.5 h-3.5 text-slate-400" />
            <span>Sinkronisasi Terakhir:</span>
            <strong className="text-slate-700">{waktuSinkronisasi || "Menunggu data..."}</strong>
          </div>
        </div>
      </header>

      {sortedKeys.length === 0 ? (
        <div className="bg-white border-2 border-dashed border-slate-300 rounded-2xl p-16 text-center flex flex-col items-center justify-center space-y-4">
          <div className="w-16 h-16 bg-sky-50 text-[#005C8A] rounded-full flex items-center justify-center">
            <Upload className="w-8 h-8" />
          </div>
          <div>
            <h3 className="text-base font-bold text-slate-800">Sistem Siap Menerima Data KOGOL</h3>
            <p className="text-xs text-slate-500 max-w-sm mx-auto mt-1">
              Silakan unggah berkas KOGOL pertama Anda melalui tombol di sudut kanan atas untuk mengaktifkan kalkulasi otomatis.
            </p>
          </div>
        </div>
      ) : (
        <>
          {/* Navigasi Tab & Pemilihan Bulan */}
          <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-3 bg-white px-5 py-3 rounded-xl border border-slate-200 shadow-xs print:hidden">
            <div className="flex items-center gap-2 flex-wrap">
              <span className="text-xs font-bold text-slate-500 uppercase tracking-wider">Periode:</span>
              <div className="flex flex-wrap gap-1.5">
                {sortedKeys.map((k) => (
                  <button
                    key={k}
                    onClick={() => {
                      setSelectedKey(k);
                      setSimulasiLunas({});
                    }}
                    className={`px-3 py-1 text-xs font-bold rounded-lg transition ${
                      k === selectedKey ? "bg-[#005C8A] text-white shadow-xs" : "bg-slate-100 text-slate-600 hover:bg-slate-200"
                    }`}
                  >
                    {dataHistory[k]?.periode || k}
                  </button>
                ))}
              </div>
            </div>

            <div className="flex bg-slate-100 p-1 rounded-lg border border-slate-200 text-xs font-semibold">
              <button
                onClick={() => setActiveTab("RINGKASAN")}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-md transition ${
                  activeTab === "RINGKASAN" ? "bg-white text-[#005C8A] shadow-xs font-bold" : "text-slate-500"
                }`}
              >
                <Building2 className="w-3.5 h-3.5" />
                <span>Dashboard Eksekutif</span>
              </button>
              <button
                onClick={() => setActiveTab("LEMBAR_KERJA")}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-md transition ${
                  activeTab === "LEMBAR_KERJA" ? "bg-white text-[#005C8A] shadow-xs font-bold" : "text-slate-500"
                }`}
              >
                <FileText className="w-3.5 h-3.5" />
                <span>Lembar Kerja Petugas</span>
              </button>
            </div>
          </div>

          {activeData && (
            <>
              {activeTab === "RINGKASAN" ? (
                <>
                  {/* Simulasi Dampak Saldo (What-If Filter) */}
                  <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-xs space-y-3 print:hidden">
                    <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-4">
                      <div className="flex items-center gap-2 text-xs">
                        <div className="p-2 rounded-lg bg-sky-50 text-[#005C8A]">
                          <Filter className="w-4 h-4" />
                        </div>
                        <div>
                          <span className="font-bold text-slate-800 uppercase tracking-wider">Simulasi Skenario Dampak Saldo</span>
                          <p className="text-slate-500">Pilih skenario atau centang pelanggan yang telah melunasi tagihan:</p>
                        </div>
                      </div>

                      <div className="flex bg-slate-100 p-1 rounded-xl border border-slate-200 text-xs font-semibold">
                        <button
                          onClick={() => setSkenario("SEMUA")}
                          className={`px-3 py-1.5 rounded-lg transition ${
                            skenario === "SEMUA" ? "bg-white text-[#005C8A] shadow-xs font-bold" : "text-slate-500"
                          }`}
                        >
                          Semua Tagihan
                        </button>
                        <button
                          onClick={() => setSkenario("TANPA_KCIC")}
                          className={`px-3 py-1.5 rounded-lg transition ${
                            skenario === "TANPA_KCIC" ? "bg-white text-[#EA580C] shadow-xs font-bold" : "text-slate-500"
                          }`}
                        >
                          Keluarkan KCIC
                        </button>
                        <button
                          onClick={() => setSkenario("HANYA_NON_AMR")}
                          className={`px-3 py-1.5 rounded-lg transition ${
                            skenario === "HANYA_NON_AMR" ? "bg-white text-emerald-700 shadow-xs font-bold" : "text-slate-500"
                          }`}
                        >
                          Hanya Non-AMR
                        </button>
                      </div>
                    </div>

                    {totalLunasSimulasi > 0 && (
                      <div className="p-2.5 bg-emerald-50 border border-emerald-200 rounded-lg flex justify-between items-center text-xs text-emerald-900">
                        <span className="flex items-center gap-2">
                          <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
                          <span>
                            Simulasi: <b>{jumlahPelangganLunasSimulasi} Pelanggan</b> ditandai lunas (Pengurangan:{" "}
                            <b>Rp {totalLunasSimulasi.toLocaleString("id-ID")}</b>).
                          </span>
                        </span>
                        <button
                          onClick={() => setSimulasiLunas({})}
                          className="text-[11px] underline font-bold hover:text-emerald-700"
                        >
                          Reset Simulasi
                        </button>
                      </div>
                    )}
                  </div>

                  {/* Target Saldo Card */}
                  <div className="bg-[#003B5C] rounded-xl p-4 text-white flex justify-between items-center shadow-xs">
                    <div>
                      <span className="text-[11px] font-bold text-sky-200 uppercase tracking-wider">
                        Pencapaian Target Saldo ({skenario.replace(/_/g, " ")}) — {activeData.periode}
                      </span>
                      <div className="text-2xl font-black text-rose-400">{capaianPersenSimulasi}%</div>
                    </div>
                    <div className="text-right text-xs text-slate-200">
                      <div>Target: <strong>Rp {(targetSaldo / 1e6).toFixed(0)} Juta</strong></div>
                      <div>
                        Sisa Tunggakan Skenario: <strong>Rp {(saldoSimulasi / 1e9).toFixed(2)} M</strong> ({plgnSimulasi} Plgn)
                      </div>
                    </div>
                  </div>

                  {/* Dual Metric Cards */}
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                    <MetricCard
                      title={`DENGAN KCIC — ${(activeData.periode || "").toUpperCase()}`}
                      totalPlgn={activeData.denganKcic?.totalPlgn || 0}
                      amrPlgn={activeData.denganKcic?.amrPlgn || 0}
                      nonAmrPlgn={activeData.denganKcic?.nonAmrPlgn || 0}
                      totalRp={activeData.denganKcic?.totalRp || 0}
                      amrRp={activeData.denganKcic?.amrRp || 0}
                      nonAmrRp={activeData.denganKcic?.nonAmrRp || 0}
                      theme="blue"
                    />
                    <MetricCard
                      title={`TANPA KCIC — ${(activeData.periode || "").toUpperCase()}`}
                      totalPlgn={activeData.tanpaKcic?.totalPlgn || 0}
                      amrPlgn={activeData.tanpaKcic?.amrPlgn || 0}
                      nonAmrPlgn={activeData.tanpaKcic?.nonAmrPlgn || 0}
                      totalRp={activeData.tanpaKcic?.totalRp || 0}
                      amrRp={activeData.tanpaKcic?.amrRp || 0}
                      nonAmrRp={activeData.tanpaKcic?.nonAmrRp || 0}
                      theme="orange"
                    />
                  </div>

                  {/* Segmentasi Berdasarkan Tarif & Daya */}
                  {activeData.rekapTarif && (
                    <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-xs">
                      <h4 className="text-xs font-bold text-slate-700 uppercase mb-3">
                        Segmentasi Piutang Berdasarkan Golongan Tarif ({activeData.periode})
                      </h4>
                      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-3">
                        {[
                          { label: "Rumah Tangga (R)", key: "R", color: "border-sky-500 text-sky-700" },
                          { label: "Bisnis (B)", key: "B", color: "border-amber-500 text-amber-700" },
                          { label: "Industri (I)", key: "I", color: "border-purple-500 text-purple-700" },
                          { label: "Publik / Kantor (P)", key: "P", color: "border-emerald-500 text-emerald-700" },
                          { label: "Kereta Cepat (T / KCIC)", key: "T", color: "border-rose-500 text-rose-700" },
                        ].map((t) => {
                          const val = activeData.rekapTarif[t.key] || { plgn: 0, rp: 0 };
                          return (
                            <div key={t.key} className={`p-3 rounded-lg border-l-4 bg-slate-50/70 border-slate-200 ${t.color}`}>
                              <span className="text-[10px] font-bold text-slate-500 block uppercase">{t.label}</span>
                              <div className="text-sm font-black text-slate-900 mt-1">
                                Rp {(val.rp / 1e6).toFixed(1)} Jt
                              </div>
                              <span className="text-[11px] text-slate-500">{val.plgn} Pelanggan</span>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  )}

                  {/* Charts Multi-Bulan & Donut (Disembunyikan saat cetak) */}
                  <div className="grid grid-cols-1 lg:grid-cols-4 gap-6 print:hidden">
                    <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-xs lg:col-span-2">
                      <h4 className="text-xs font-bold text-slate-700 uppercase mb-4">
                        Tingkat Realisasi Multi-Bulan (Dengan vs Tanpa KCIC)
                      </h4>
                      <div className="h-60">
                        <ResponsiveContainer width="100%" height="100%">
                          <BarChart data={barChartData}>
                            <CartesianGrid strokeDasharray="3 3" vertical={false} />
                            <XAxis dataKey="bulan" />
                            <YAxis tickFormatter={(val) => `${(val / 1e9).toFixed(1)}M`} />
                            <Tooltip formatter={(val: any) => `Rp ${Number(val).toLocaleString("id-ID")}`} />
                            <Bar dataKey="Dengan_KCIC" fill="#EA580C" radius={[4, 4, 0, 0]} name="Dengan KCIC" maxBarSize={45} />
                            <Bar dataKey="Tanpa_KCIC" fill="#005C8A" radius={[4, 4, 0, 0]} name="Tanpa KCIC" maxBarSize={45} />
                          </BarChart>
                        </ResponsiveContainer>
                      </div>
                    </div>

                    <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-xs">
                      <h4 className="text-xs font-bold text-slate-700 uppercase mb-4">
                        % Pelanggan ({(activeData.periode || "").split(" ")[0]})
                      </h4>
                      <div className="h-60">
                        <ResponsiveContainer width="100%" height="100%">
                          <PieChart>
                            <Pie data={donutPelanggan} dataKey="value" innerRadius={50} outerRadius={75} paddingAngle={4}>
                              {donutPelanggan.map((entry, idx) => (
                                <Cell key={`cell-${idx}`} fill={entry.color} />
                              ))}
                            </Pie>
                            <Tooltip />
                          </PieChart>
                        </ResponsiveContainer>
                      </div>
                    </div>

                    <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-xs">
                      <h4 className="text-xs font-bold text-slate-700 uppercase mb-4">
                        % Rupiah ({(activeData.periode || "").split(" ")[0]})
                      </h4>
                      <div className="h-60">
                        <ResponsiveContainer width="100%" height="100%">
                          <PieChart>
                            <Pie data={donutRupiah} dataKey="value" innerRadius={50} outerRadius={75} paddingAngle={4}>
                              {donutRupiah.map((entry, idx) => (
                                <Cell key={`cell-rp-${idx}`} fill={entry.color} />
                              ))}
                            </Pie>
                            <Tooltip formatter={(val: any) => `Rp ${Number(val).toLocaleString("id-ID")}`} />
                          </PieChart>
                        </ResponsiveContainer>
                      </div>
                    </div>
                  </div>

                  {/* Panel Komparasi Antar-Bulan */}
                  {hasComparison && prevData && currData && (
                    <div className="bg-white rounded-xl border border-slate-200 shadow-xs overflow-hidden print:border-none">
                      <div className="bg-gradient-to-r from-slate-900 to-[#003B5C] px-6 py-4 text-white flex flex-col md:flex-row justify-between md:items-center gap-2 print:bg-white print:text-black print:p-0 print:border-b">
                        <div>
                          <h3 className="text-sm font-black tracking-wider uppercase flex items-center gap-2">
                            <span>📊</span> Analisis Komparasi Kinerja: {prevData.periode} vs {currData.periode}
                          </h3>
                          <p className="text-xs text-sky-200 print:text-slate-600">
                            Evaluasi kenaikan / penurunan saldo piutang tunggakan antar-bulan secara otomatis
                          </p>
                        </div>
                        <div className="flex items-center gap-2 text-xs bg-white/10 px-3 py-1.5 rounded-lg backdrop-blur-xs font-semibold print:hidden">
                          <span>{(prevData.periode || "").split(" ")[0]}</span>
                          <ArrowRight className="w-3.5 h-3.5" />
                          <span className="text-[#FFD100]">{(currData.periode || "").split(" ")[0]}</span>
                        </div>
                      </div>

                      <div className="p-6 grid grid-cols-1 md:grid-cols-3 gap-4 border-b border-slate-100 bg-slate-50/50 print:bg-white print:p-2">
                        <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-xs">
                          <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider">
                            Selisih Total Saldo (Dengan KCIC)
                          </span>
                          <div className="flex items-baseline gap-2 mt-2">
                            <span className={`text-xl font-black ${diffTotalRp > 0 ? "text-rose-600" : "text-emerald-600"}`}>
                              {diffTotalRp > 0 ? "+" : ""}Rp {Math.abs(diffTotalRp).toLocaleString("id-ID")}
                            </span>
                          </div>
                          <div className="flex items-center gap-1.5 mt-2">
                            {diffTotalRp > 0 ? (
                              <span className="inline-flex items-center gap-1 text-[11px] font-bold px-2 py-0.5 rounded bg-rose-50 text-rose-600">
                                <TrendingUp className="w-3.5 h-3.5" /> +{pctTotalRp.toFixed(1)}% (Memburuk/Naik)
                              </span>
                            ) : (
                              <span className="inline-flex items-center gap-1 text-[11px] font-bold px-2 py-0.5 rounded bg-emerald-50 text-emerald-600">
                                <TrendingDown className="w-3.5 h-3.5" /> {pctTotalRp.toFixed(1)}% (Membaik/Turun)
                              </span>
                            )}
                          </div>
                        </div>

                        <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-xs">
                          <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider">
                            Selisih Saldo Reguler (Tanpa KCIC)
                          </span>
                          <div className="flex items-baseline gap-2 mt-2">
                            <span className={`text-xl font-black ${diffTanpaKcicRp > 0 ? "text-rose-600" : "text-emerald-600"}`}>
                              {diffTanpaKcicRp > 0 ? "+" : ""}Rp {Math.abs(diffTanpaKcicRp).toLocaleString("id-ID")}
                            </span>
                          </div>
                          <div className="flex items-center gap-1.5 mt-2">
                            {diffTanpaKcicRp > 0 ? (
                              <span className="inline-flex items-center gap-1 text-[11px] font-bold px-2 py-0.5 rounded bg-rose-50 text-rose-600">
                                <TrendingUp className="w-3.5 h-3.5" /> +{pctTanpaKcicRp.toFixed(1)}% (Naik)
                              </span>
                            ) : (
                              <span className="inline-flex items-center gap-1 text-[11px] font-bold px-2 py-0.5 rounded bg-emerald-50 text-emerald-600">
                                <TrendingDown className="w-3.5 h-3.5" /> {pctTanpaKcicRp.toFixed(1)}% (Terkendali)
                              </span>
                            )}
                          </div>
                        </div>

                        <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-xs">
                          <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider">
                            Perubahan Jumlah Pelanggan
                          </span>
                          <div className="flex items-baseline gap-2 mt-2">
                            <span className={`text-xl font-black ${diffPlgn > 0 ? "text-amber-600" : "text-emerald-600"}`}>
                              {diffPlgn > 0 ? `+${diffPlgn}` : diffPlgn} Plgn
                            </span>
                          </div>
                          <div className="text-xs text-slate-500 mt-2">
                            Dari <b>{prevData.denganKcic?.totalPlgn || 0}</b> menjadi <b>{currData.denganKcic?.totalPlgn || 0}</b> pelanggan
                          </div>
                        </div>
                      </div>

                      <div className="p-6 space-y-4 print:p-2">
                        <div className="flex items-center justify-between border-b border-slate-100 pb-2">
                          <div className="flex items-center gap-2">
                            <div className="w-2 h-4 bg-[#005C8A] rounded-full"></div>
                            <h4 className="text-xs font-extrabold text-slate-800 uppercase tracking-wider">
                              Executive Evaluation & Actionable Insight ({(prevData.periode || "").split(" ")[0]} vs {(currData.periode || "").split(" ")[0]})
                            </h4>
                          </div>
                          <span className="text-[11px] font-semibold text-slate-400">
                            Unit Pelaksana Pelayanan Pelanggan (UP3) Jatinegara
                          </span>
                        </div>

                        {diffTotalRp > 0 ? (
                          <div className="bg-rose-50/80 border border-rose-200 rounded-xl p-4 flex flex-col md:flex-row items-start md:items-center justify-between gap-3">
                            <div className="flex items-start gap-3">
                              <div className="p-2 bg-rose-100 text-rose-700 rounded-lg shrink-0 mt-0.5">
                                <AlertCircle className="w-5 h-5" />
                              </div>
                              <div>
                                <h5 className="text-xs font-black text-rose-950 uppercase tracking-wide">
                                  Peringatan: Kenaikan Saldo Tunggakan Rekening
                                </h5>
                                <p className="text-xs text-rose-800 mt-0.5 leading-relaxed">
                                  Saldo akhir pada periode <b>{currData.periode}</b> meningkat sebesar{" "}
                                  <b>Rp {diffTotalRp.toLocaleString("id-ID")}</b> (+{pctTotalRp.toFixed(1)}%) dibanding {prevData.periode}.
                                </p>
                              </div>
                            </div>
                            <div className="bg-white px-3 py-1.5 rounded-lg border border-rose-200 text-right shrink-0">
                              <span className="text-[10px] uppercase font-bold text-rose-500 block">Deviasi Saldo</span>
                              <span className="text-sm font-black text-rose-700">+{pctTotalRp.toFixed(1)}% MoM</span>
                            </div>
                          </div>
                        ) : (
                          <div className="bg-emerald-50/80 border border-emerald-200 rounded-xl p-4 flex flex-col md:flex-row items-start md:items-center justify-between gap-3">
                            <div className="flex items-start gap-3">
                              <div className="p-2 bg-emerald-100 text-emerald-700 rounded-lg shrink-0 mt-0.5">
                                <CheckCircle2 className="w-5 h-5" />
                              </div>
                              <div>
                                <h5 className="text-xs font-black text-emerald-950 uppercase tracking-wide">
                                  Kinerja Positif: Penurunan Saldo Tunggakan Efektif
                                </h5>
                                <p className="text-xs text-emerald-800 mt-0.5 leading-relaxed">
                                  Realisasi saldo akhir berhasil ditekan sebesar{" "}
                                  <b>Rp {Math.abs(diffTotalRp).toLocaleString("id-ID")}</b> ({pctTotalRp.toFixed(1)}%) dibanding {prevData.periode}.
                                </p>
                              </div>
                            </div>
                            <div className="bg-white px-3 py-1.5 rounded-lg border border-emerald-200 text-right shrink-0">
                              <span className="text-[10px] uppercase font-bold text-emerald-500 block">Efisiensi Saldo</span>
                              <span className="text-sm font-black text-emerald-700">{pctTotalRp.toFixed(1)}% MoM</span>
                            </div>
                          </div>
                        )}

                        <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                          <div className="bg-slate-50 border border-slate-200 rounded-xl p-3.5 space-y-1.5">
                            <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1.5">
                              <span>🚆</span> Pengaruh Portofolio KCIC
                            </span>
                            <div className="text-xs text-slate-800 leading-relaxed">
                              {diffKcicRp > 0 ? (
                                <>
                                  Saldo KCIC bertambah <b>Rp {diffKcicRp.toLocaleString("id-ID")}</b>. Kenaikan ini menyumbang{" "}
                                  <b>{kcicContributionPct}%</b> dari seluruh lonjakan tunggakan unit.
                                </>
                              ) : diffKcicRp < 0 ? (
                                <>
                                  Penagihan KCIC berhasil menurunkan saldo sebesar{" "}
                                  <b>Rp {Math.abs(diffKcicRp).toLocaleString("id-ID")}</b>.
                                </>
                              ) : (
                                "Saldo tunggakan segmen KCIC terpantau stagnan tanpa mutasi nominal."
                              )}
                            </div>
                          </div>

                          <div className="bg-slate-50 border border-slate-200 rounded-xl p-3.5 space-y-1.5">
                            <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1.5">
                              <span>⚡</span> Saldo Murni Reguler (Tanpa KCIC)
                            </span>
                            <div className="text-xs text-slate-800 leading-relaxed">
                              {diffTanpaKcicRp > 0 ? (
                                <>
                                  Tunggakan pelanggan reguler mengalami kenaikan sebesar{" "}
                                  <b className="text-rose-600">Rp {diffTanpaKcicRp.toLocaleString("id-ID")}</b> (+{pctTanpaKcicRp.toFixed(1)}%).
                                </>
                              ) : (
                                <>
                                  Tunggakan reguler terpantau membaik dengan reduksi sebesar{" "}
                                  <b className="text-emerald-600">Rp {Math.abs(diffTanpaKcicRp).toLocaleString("id-ID")}</b> ({pctTanpaKcicRp.toFixed(1)}%).
                                </>
                              )}
                            </div>
                          </div>

                          <div className="bg-slate-50 border border-slate-200 rounded-xl p-3.5 space-y-1.5">
                            <span className="text-[10px] font-bold text-[#005C8A] uppercase tracking-wider flex items-center gap-1.5">
                              <span>📋</span> Rekomendasi Tindak Lanjut
                            </span>
                            <div className="text-xs text-slate-700 leading-relaxed">
                              {diffTotalRp > 0
                                ? "Prioritaskan rekonsiliasi faktur tagihan KCIC terpusat serta lakukan penertiban surat pemutusan sementara untuk IDPEL AMR reguler berpiutang tinggi."
                                : "Pertahankan ritme penagihan harian siklus awal dan monitor pengawasan pembongkaran rampung pelanggan KOGOL yang belum melunasi."}
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>
                  )}
                </>
              ) : null}

              {/* TABEL PELANGGAN PARETO / LEMBAR KERJA */}
              <div className="bg-white rounded-xl border border-slate-200 shadow-xs overflow-hidden print:border-none">
                <div className="bg-slate-900 px-6 py-4 text-white flex flex-col md:flex-row justify-between md:items-center gap-3 print:bg-white print:text-black print:p-0 print:border-b">
                  <div className="flex items-center gap-3">
                    <div className="p-2 rounded-lg bg-amber-500/20 text-amber-400 print:hidden">
                      <Flame className="w-5 h-5" />
                    </div>
                    <div>
                      <h3 className="text-sm font-black tracking-wider uppercase">
                        {activeTab === "LEMBAR_KERJA"
                          ? `Lembar Kerja Penagihan Lapangan (Tusbung / Billman) — ${activeData.periode}`
                          : `Prioritas Penagihan: Top Pelanggan Tunggakan Terbesar — ${activeData.periode}`}
                      </h3>
                      <p className="text-xs text-slate-400 print:text-slate-600">
                        {activeTab === "LEMBAR_KERJA"
                          ? "Dokumen instruksi kerja penagihan lapangan, pembongkaran rampung, dan penertiban sambungan tenaga listrik."
                          : "Gunakan centang kotak bayar di samping baris untuk mensimulasikan dampak saldo jika pelanggan melunasi tagihan."}
                      </p>
                    </div>
                  </div>

                  <div className="flex flex-wrap items-center gap-2 print:hidden">
                    <div className="relative">
                      <Search className="w-3.5 h-3.5 absolute left-3 top-2.5 text-slate-400" />
                      <input
                        type="text"
                        placeholder="Cari IDPEL / Nama..."
                        value={searchQuery}
                        onChange={(e) => setSearchQuery(e.target.value)}
                        className="pl-8 pr-3 py-1.5 bg-slate-800 border border-slate-700 rounded-lg text-xs text-white placeholder-slate-400 focus:outline-none focus:border-sky-400 w-44"
                      />
                    </div>

                    <select
                      value={filterKategori}
                      onChange={(e) => setFilterKategori(e.target.value)}
                      className="bg-slate-800 border border-slate-700 text-xs text-slate-200 px-2.5 py-1.5 rounded-lg focus:outline-none focus:border-sky-400"
                    >
                      <option value="SEMUA">Semua Kategori</option>
                      <option value="KCIC">Khusus KCIC</option>
                      <option value="AMR">AMR Reguler</option>
                      <option value="NON-AMR">Non-AMR</option>
                    </select>
                  </div>
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full text-xs text-left">
                    <thead className="bg-slate-50 text-slate-600 font-bold border-b border-slate-200 print:bg-slate-100">
                      <tr>
                        <th className="py-2.5 px-3 w-10 text-center print:hidden">Bayar</th>
                        <th className="py-2.5 px-3 w-10 text-center">No</th>
                        <th className="py-2.5 px-4">ID Pelanggan</th>
                        <th className="py-2.5 px-4">Nama Pelanggan</th>
                        <th className="py-2.5 px-4">Tarif / Daya</th>
                        <th className="py-2.5 px-4 text-center">Segmen</th>
                        <th className="py-2.5 px-4 text-right">Rupiah Tunggakan</th>
                        {activeTab === "LEMBAR_KERJA" && (
                          <>
                            <th className="py-2.5 px-4 text-center w-28">Status Lapangan</th>
                            <th className="py-2.5 px-4 text-center w-28">Paraf Pelanggan</th>
                          </>
                        )}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 text-slate-700 print:divide-slate-300">
                      {filteredPelanggan.length > 0 ? (
                        filteredPelanggan.slice(0, activeTab === "RINGKASAN" ? 15 : 60).map((item, idx) => {
                          const isLunas = !!simulasiLunas[item.idpel];
                          return (
                            <tr
                              key={item.idpel}
                              className={`transition ${
                                isLunas ? "bg-emerald-50/70 line-through text-slate-400" : "hover:bg-slate-50"
                              }`}
                            >
                              <td className="py-2 px-3 text-center print:hidden">
                                <button
                                  onClick={() => togglePelunasan(item.idpel)}
                                  className="text-slate-400 hover:text-emerald-600 transition"
                                  title="Tandai Sudah Bayar / Simulasi"
                                >
                                  {isLunas ? (
                                    <CheckSquare className="w-4 h-4 text-emerald-600" />
                                  ) : (
                                    <Square className="w-4 h-4" />
                                  )}
                                </button>
                              </td>
                              <td className="py-2 px-3 text-center font-bold text-slate-400">{idx + 1}</td>
                              <td className="py-2 px-4 font-mono font-semibold text-slate-900">{item.idpel}</td>
                              <td className="py-2 px-4 font-bold text-slate-800">{item.nama}</td>
                              <td className="py-2 px-4 text-slate-600">{item.tarifDaya}</td>
                              <td className="py-2 px-4 text-center">
                                {item.kategori === "KCIC" && (
                                  <span className="px-2 py-0.5 text-[10px] font-extrabold rounded bg-rose-100 text-rose-700 border border-rose-200">
                                    KCIC
                                  </span>
                                )}
                                {item.kategori === "AMR" && (
                                  <span className="px-2 py-0.5 text-[10px] font-extrabold rounded bg-sky-100 text-[#005C8A] border border-sky-200">
                                    AMR
                                  </span>
                                )}
                                {item.kategori === "NON-AMR" && (
                                  <span className="px-2 py-0.5 text-[10px] font-extrabold rounded bg-amber-100 text-amber-800 border border-amber-200">
                                    NON-AMR
                                  </span>
                                )}
                              </td>
                              <td className="py-2 px-4 text-right font-black text-slate-900">
                                Rp {item.rpTunggakan.toLocaleString("id-ID")}
                              </td>

                              {activeTab === "LEMBAR_KERJA" && (
                                <>
                                  <td className="py-2 px-4 text-center border-l border-slate-200 text-[10px] text-slate-500">
                                    [ ] Lunas &nbsp; [ ] Segel
                                  </td>
                                  <td className="py-2 px-4 text-center border-l border-slate-200 text-slate-300">
                                    ...................
                                  </td>
                                </>
                              )}
                            </tr>
                          );
                        })
                      ) : (
                        <tr>
                          <td colSpan={8} className="py-6 text-center text-slate-400">
                            Tidak ditemukan data pelanggan yang sesuai dengan kriteria pencarian.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>

                {/* Kolom Pengesahan Tiga Pejabat (Hanya Muncul Saat Dicetak ke PDF/Kertas) */}
                <div className="hidden print:grid grid-cols-3 gap-6 p-6 mt-8 text-center text-xs text-slate-900 border-t-2 border-slate-800">
                  <div>
                    <p className="font-semibold text-slate-600">Disusun oleh,</p>
                    <p className="font-bold text-slate-900 mt-1">Team Leader Pengelolaan Piutang</p>
                    <div className="h-20"></div>
                    <p className="font-bold underline text-slate-900">( ............................................ )</p>
                    <p className="text-[10px] text-slate-500 mt-0.5">NIP: .......................................</p>
                  </div>
                  <div>
                    <p className="font-semibold text-slate-600">Diperiksa oleh,</p>
                    <p className="font-bold text-slate-900 mt-1">Asisten Manajer Transaksi Energi</p>
                    <div className="h-20"></div>
                    <p className="font-bold underline text-slate-900">( ............................................ )</p>
                    <p className="text-[10px] text-slate-500 mt-0.5">NIP: .......................................</p>
                  </div>
                  <div>
                    <p className="font-semibold text-slate-600">Disetujui oleh,</p>
                    <p className="font-bold text-slate-900 mt-1">Manager UP3 Jatinegara</p>
                    <div className="h-20"></div>
                    <p className="font-bold underline text-slate-900">( ............................................ )</p>
                    <p className="text-[10px] text-slate-500 mt-0.5">NIP: .......................................</p>
                  </div>
                </div>
              </div>
            </>
          )}
        </>
      )}
    </main>
  );
}