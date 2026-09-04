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
  Search,
  Printer,
  CheckSquare,
  Square,
  FileText,
  Building2,
  Users,
  ArrowUpDown,
  Trophy,
  Target,
  Layers,
  BarChart3,
} from "lucide-react";
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

const PlnLogo = ({ className = "w-8 h-10" }: { className?: string }) => (
  <svg viewBox="0 0 100 125" className={className} fill="none" xmlns="http://www.w3.org/2000/svg">
    <path d="M10 8 H90 L80 115 L20 115 Z" fill="#FFDD00" stroke="#F59E0B" strokeWidth="2" />
    <path d="M24 78 C32 74, 40 82, 48 78 C56 74, 64 82, 74 78" stroke="#005C8A" strokeWidth="4.5" strokeLinecap="round" fill="none" />
    <path d="M23 88 C31 84, 39 92, 47 88 C55 84, 63 92, 75 88" stroke="#005C8A" strokeWidth="4.5" strokeLinecap="round" fill="none" />
    <path d="M22 98 C30 94, 38 102, 46 98 C54 94, 62 102, 76 98" stroke="#005C8A" strokeWidth="4.5" strokeLinecap="round" fill="none" />
    <path d="M58 14 L32 54 L52 54 L42 82 L72 42 L52 42 Z" fill="#E11D48" stroke="#BE123C" strokeWidth="1.5" strokeLinejoin="round" />
  </svg>
);

type SkenarioSimulasi = "SEMUA" | "TANPA_KCIC" | "HANYA_NON_AMR";
type KolomSortir = "NOMINAL" | "REGU" | "PETUGAS" | "NAMA";

export default function DashboardPage() {
  const [dataHistory, setDataHistory] = useState<Record<string, HasilKogol>>({});
  const [selectedKey, setSelectedKey] = useState<string>("");
  const [modeLaporan] = useState<"BULANAN" | "TAHUNAN">("BULANAN");
  const [waktuSinkronisasi, setWaktuSinkronisasi] = useState<string>("");
  const [skenario, setSkenario] = useState<SkenarioSimulasi>("SEMUA");
  const [searchQuery, setSearchQuery] = useState<string>("");
  const [filterKategori, setFilterKategori] = useState<string>("SEMUA");
  const [filterRegu, setFilterRegu] = useState<string>("SEMUA");
  const [filterPetugas, setFilterPetugas] = useState<string>("SEMUA");
  const [kolomSortir, setKolomSortir] = useState<KolomSortir>("NOMINAL");
  const [arahSortir, setArahSortir] = useState<"ASC" | "DESC">("DESC");
  const [simulasiLunas, setSimulasiLunas] = useState<Record<string, boolean>>({});
  const [activeTab, setActiveTab] = useState<"RINGKASAN" | "LEMBAR_KERJA">("RINGKASAN");

  const fileInputRef = useRef<HTMLInputElement>(null);
  const targetSaldo = 982878731;

  useEffect(() => {
    const now = new Date();
    const formatted = new Intl.DateTimeFormat("id-ID", {
      day: "2-digit",
      month: "short",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit",
      hour12: false,
    }).format(now);
    setWaktuSinkronisasi(`${formatted} WIB`);

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

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    const readFileBuffer = (file: File): Promise<{ buffer: ArrayBuffer; name: string }> => {
      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (evt) => {
          if (evt.target?.result) {
            resolve({ buffer: evt.target.result as ArrayBuffer, name: file.name });
          } else {
            reject(new Error(`Gagal membaca berkas ${file.name}`));
          }
        };
        reader.onerror = () => reject(new Error(`Error membaca berkas ${file.name}`));
        reader.readAsArrayBuffer(file);
      });
    };

    try {
      const fileBuffers = await Promise.all(
        Array.from(files).map((file) => readFileBuffer(file))
      );

      const parsedResults: Record<string, HasilKogol> = {};
      let lastKey = "";

      fileBuffers.forEach(({ buffer, name }) => {
        try {
          const parsed = parseExcelKogolBuffer(buffer, name || "file.xlsx");
          const key = `${parsed.tahun}-${parsed.kodeBulan}`;
          parsedResults[key] = parsed;
          lastKey = key;
        } catch (parseErr) {
          console.error(`Gagal mengurai berkas ${name}:`, parseErr);
        }
      });

      if (Object.keys(parsedResults).length === 0) {
        alert("Tidak ada berkas Excel KOGOL valid yang berhasil dibaca.");
        return;
      }

      setDataHistory((prev) => {
        const updated = { ...prev, ...parsedResults };
        try {
          localStorage.setItem("satu_jatinegara_history", JSON.stringify(updated));
        } catch (storageErr) {
          console.warn("Penyimpanan browser penuh:", storageErr);
        }
        return updated;
      });

      if (lastKey) {
        setSelectedKey(lastKey);
      }
      setSimulasiLunas({});
      setFilterRegu("SEMUA");
      setFilterPetugas("SEMUA");

      const now = new Date();
      const formatted = new Intl.DateTimeFormat("id-ID", {
        day: "2-digit",
        month: "short",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
        hour12: false,
      }).format(now);
      setWaktuSinkronisasi(`${formatted} WIB`);
    } catch (err) {
      console.error("Gagal memproses berkas yang dipilih:", err);
      alert("Gagal membaca struktur berkas Excel.");
    } finally {
      if (fileInputRef.current) fileInputRef.current.value = "";
    }
  };

  const handleResetData = () => {
    localStorage.removeItem("satu_jatinegara_history");
    setDataHistory({});
    setSelectedKey("");
    setSimulasiLunas({});
    setFilterRegu("SEMUA");
    setFilterPetugas("SEMUA");
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

  const capaianPersenSimulasi = useMemo(() => {
    if (!activeData || sortedKeys.length === 0) return "0.0";

    const currentIdx = sortedKeys.indexOf(selectedKey);
    const slicedKeys = sortedKeys.slice(0, currentIdx !== -1 ? currentIdx + 1 : sortedKeys.length);
    const pembagi = Math.max(1, slicedKeys.length);

    let totalKumulatif = 0;
    slicedKeys.forEach((k) => {
      if (k === selectedKey) {
        totalKumulatif += saldoSimulasi;
      } else {
        totalKumulatif += dataHistory[k]?.denganKcic?.totalRp || 0;
      }
    });

    const rataKumulatif = totalKumulatif / pembagi;
    return ((2 - (rataKumulatif / targetSaldo)) * 100).toFixed(1);
  }, [activeData, sortedKeys, selectedKey, saldoSimulasi, dataHistory, targetSaldo]);

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

  const opsiPetugas = useMemo(() => {
    if (!activeData?.daftarPetugas) return [];
    if (filterRegu === "SEMUA") return activeData.daftarPetugas;
    return activeData.daftarPetugas.filter((p) => p.regu === filterRegu);
  }, [activeData, filterRegu]);

  const filteredPelanggan = useMemo(() => {
    if (!activeData || !activeData.semuaPelanggan) return [];

    const hasil = activeData.semuaPelanggan.filter((p) => {
      const matchText =
        (p.idpel || "").toLowerCase().includes(searchQuery.toLowerCase()) ||
        (p.nama || "").toLowerCase().includes(searchQuery.toLowerCase()) ||
        (p.rbm || "").toLowerCase().includes(searchQuery.toLowerCase());

      const matchKategori = filterKategori === "SEMUA" || p.kategori === filterKategori;
      const matchRegu = filterRegu === "SEMUA" || p.regu === filterRegu;
      const matchPetugas =
        filterPetugas === "SEMUA" ||
        p.idPetugas === filterPetugas ||
        (p as any).kodeAnggota === filterPetugas ||
        (p as any).kodePerorangan === filterPetugas;

      return matchText && matchKategori && matchRegu && matchPetugas;
    });

    hasil.sort((a, b) => {
      const multiplier = arahSortir === "ASC" ? 1 : -1;
      if (kolomSortir === "NOMINAL") {
        return ((a.rpTunggakan || 0) - (b.rpTunggakan || 0)) * multiplier;
      } else if (kolomSortir === "REGU") {
        return String(a.regu || "-").localeCompare(String(b.regu || "-")) * multiplier;
      } else if (kolomSortir === "PETUGAS") {
        const ptgA = String((a as any).kodeAnggota || a.idPetugas || "-");
        const ptgB = String((b as any).kodeAnggota || b.idPetugas || "-");
        return ptgA.localeCompare(ptgB) * multiplier;
      } else {
        return String(a.nama || "").localeCompare(String(b.nama || "")) * multiplier;
      }
    });

    return hasil;
  }, [activeData, searchQuery, filterKategori, filterRegu, filterPetugas, kolomSortir, arahSortir]);

  const handleToggleSort = (kolom: KolomSortir) => {
    if (kolomSortir === kolom) {
      setArahSortir((prev) => (prev === "ASC" ? "DESC" : "ASC"));
    } else {
      setKolomSortir(kolom);
      setArahSortir("ASC");
    }
  };

  const peringkatRegu = useMemo(() => {
    if (!activeData || !activeData.semuaPelanggan) {
      return { list: [] as { regu: string; totalRp: number; totalPlgn: number }[], totalRp: 0, totalPlgn: 0 };
    }

    const map = new Map<string, { regu: string; totalRp: number; totalPlgn: number }>();

    activeData.semuaPelanggan.forEach((p) => {
      if (!p.regu || p.regu === "-") return;
      const existing = map.get(p.regu) || { regu: p.regu, totalRp: 0, totalPlgn: 0 };
      existing.totalRp += p.rpTunggakan || 0;
      existing.totalPlgn += 1;
      map.set(p.regu, existing);
    });

    const list = Array.from(map.values()).sort((a, b) => b.totalRp - a.totalRp);
    const totalRp = list.reduce((acc, r) => acc + r.totalRp, 0);
    const totalPlgn = list.reduce((acc, r) => acc + r.totalPlgn, 0);

    return { list, totalRp, totalPlgn };
  }, [activeData]);

  const perbandinganBulanan = useMemo(() => {
    let kumulatifDenganKcic = 0;

    const list = sortedKeys.map((k, idx) => {
      const item = dataHistory[k];
      const totalRp = item?.denganKcic?.totalRp || 0;
      const totalPlgn = item?.denganKcic?.totalPlgn || 0;

      kumulatifDenganKcic += totalRp;
      const pembagiBulan = idx + 1;
      const rataKumulatifDenganKcic = kumulatifDenganKcic / pembagiBulan;

      const persenPolaritas = (2 - (rataKumulatifDenganKcic / targetSaldo)) * 100;

      return {
        key: k,
        periode: item?.periode || k,
        totalRp,
        totalPlgn,
        pembagiBulan,
        persen: persenPolaritas,
      };
    });

    const totalRpGabungan = list.reduce((acc, x) => acc + x.totalRp, 0);
    const totalPlgnGabungan = list.reduce((acc, x) => acc + x.totalPlgn, 0);
    const jumlahBulan = list.length;

    const rataRataRp = jumlahBulan > 0 ? totalRpGabungan / jumlahBulan : 0;
    const rataRataPlgn = jumlahBulan > 0 ? Math.round(totalPlgnGabungan / jumlahBulan) : 0;
    const persenTotalAkhir = jumlahBulan > 0 ? (2 - (rataRataRp / targetSaldo)) * 100 : 0;

    return {
      list,
      totalRpGabungan,
      totalPlgnGabungan,
      jumlahBulan,
      rataRataRp,
      rataRataPlgn,
      persenTotalAkhir,
    };
  }, [sortedKeys, dataHistory, targetSaldo]);

  const barChartData = sortedKeys
    .map((k) => {
      const item = dataHistory[k];
      if (!item) return null;
      const namaBulanSingkat = (item.periode || "").split(" ")[0]?.substring(0, 3) || k;
      return {
        bulan: `${namaBulanSingkat} '${(item.tahun || "26").substring(2)}`,
        Dengan_KCIC: item.denganKcic?.totalRp || 0,
        Tanpa_KCIC: item.tanpaKcic?.totalRp || 0,
      };
    })
    .filter(Boolean);

  const handleDownloadExcel = () => {
    if (!activeData || sortedKeys.length === 0) return;

    const wb = XLSX.utils.book_new();

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
      ["Pencapaian Terhadap Target (Polaritas Negatif)", `${capaianPersenSimulasi}%`],
      [],
      ["REKAPITULASI REALISASI SALDO MULTI-PERIODE (RUPIAH)"],
      ["Kategori", "Jenis Pelanggan", ...sortedKeys.map((k) => dataHistory[k].periode)],
      ["Dengan KCIC", "AMR", ...sortedKeys.map((k) => dataHistory[k].denganKcic.amrRp)],
      ["Dengan KCIC", "NON-AMR", ...sortedKeys.map((k) => dataHistory[k].denganKcic.nonAmrRp)],
      ["Tanpa KCIC", "AMR", ...sortedKeys.map((k) => dataHistory[k].tanpaKcic.amrRp)],
      ["Tanpa KCIC", "NON-AMR", ...sortedKeys.map((k) => dataHistory[k].tanpaKcic.nonAmrRp)],
      ["Khusus KCIC", "AMR KCIC", ...sortedKeys.map((k) => dataHistory[k].kcicOnly.rp)],
      [],
      ["PERINGKAT REGU BERDASARKAN TOTAL TUNGGAKAN"],
      ["Peringkat", "Regu", "Jumlah Pelanggan", "Total Tunggakan (Rp)", "Kontribusi (%)"],
      ...peringkatRegu.list.map((r, idx) => [
        idx + 1,
        r.regu,
        r.totalPlgn,
        r.totalRp,
        `${peringkatRegu.totalRp > 0 ? ((r.totalRp / peringkatRegu.totalRp) * 100).toFixed(2) : "0.00"}%`,
      ]),
      ["", "TOTAL", peringkatRegu.totalPlgn, peringkatRegu.totalRp, "100.00%"],
    ];

    const ws1 = XLSX.utils.aoa_to_sheet(sheet1Data);
    ws1["!cols"] = [{ wch: 34 }, { wch: 22 }, { wch: 24 }, { wch: 24 }, { wch: 24 }];
    XLSX.utils.book_append_sheet(wb, ws1, "Ringkasan Eksekutif");

    const top10Data: any[][] = [
      ["TOP 10 PELANGGAN SALDO TUNGGAKAN TERBESAR (PARETO 80/20)"],
      [`PERIODE: ${activeData.periode.toUpperCase()}`],
      [],
      ["No", "ID Pelanggan", "Nama Pelanggan", "Tarif / Daya", "No RBM", "Regu", "Kode Anggota", "Segmen", "Nominal Tunggakan (Rp)", "Kontribusi (%)"],
    ];

    const top10List = (activeData.semuaPelanggan || []).slice(0, 10);
    top10List.forEach((p, idx) => {
      const kontribusi = ((p.rpTunggakan / (activeData.denganKcic.totalRp || 1)) * 100).toFixed(2);
      top10Data.push([
        idx + 1,
        p.idpel,
        p.nama,
        p.tarifDaya,
        p.rbm,
        p.regu,
        (p as any).kodeAnggota || p.idPetugas,
        p.kategori,
        p.rpTunggakan,
        `${kontribusi}%`,
      ]);
    });

    const ws2 = XLSX.utils.aoa_to_sheet(top10Data);
    ws2["!cols"] = [{ wch: 6 }, { wch: 18 }, { wch: 35 }, { wch: 16 }, { wch: 16 }, { wch: 10 }, { wch: 14 }, { wch: 12 }, { wch: 24 }, { wch: 15 }];
    XLSX.utils.book_append_sheet(wb, ws2, "Top 10 Prioritas");

    const bandingBulananData: any[][] = [
      ["PERBANDINGAN & RATA-RATA TAGIHAN ANTAR BULAN"],
      [`TARGET SALDO RESMI: Rp ${targetSaldo.toLocaleString("id-ID")}`],
      ["FORMULA POLARITAS NEGATIF: 2 - (Realisasi Kumulatif Dengan KCIC / Target)"],
      [],
      ["Periode", "Jumlah Pelanggan", "Total Tagihan (Dengan KCIC)", "Pencapaian Target (%)"],
    ];
    perbandinganBulanan.list.forEach((b) => {
      bandingBulananData.push([
        b.periode,
        b.totalPlgn,
        b.totalRp,
        `${b.persen.toFixed(2)}%`,
      ]);
    });
    bandingBulananData.push([
      `TOTAL GABUNGAN (${perbandinganBulanan.jumlahBulan} Bulan)`,
      perbandinganBulanan.totalPlgnGabungan,
      perbandinganBulanan.totalRpGabungan,
      "—",
    ]);
    bandingBulananData.push([
      "RATA-RATA / BULAN",
      perbandinganBulanan.rataRataPlgn,
      Math.round(perbandinganBulanan.rataRataRp),
      `${perbandinganBulanan.persenTotalAkhir.toFixed(2)}%`,
    ]);

    const wsBanding = XLSX.utils.aoa_to_sheet(bandingBulananData);
    wsBanding["!cols"] = [{ wch: 20 }, { wch: 18 }, { wch: 28 }, { wch: 25 }];
    XLSX.utils.book_append_sheet(wb, wsBanding, "Perbandingan Bulanan");

    const semuaData: any[][] = [
      ["DAFTAR LENGKAP TUNGGAKAN PELANGGAN KOGOL UNIT JATINEGARA"],
      [`PERIODE: ${activeData.periode.toUpperCase()}`],
      [`FILTER AKTIF: Regu [${filterRegu}] | Petugas [${filterPetugas}] | Kategori [${filterKategori}]`],
      [],
      ["No", "ID Pelanggan", "Nama Pelanggan", "Tarif / Daya", "Nomor RBM", "Regu", "Kode Anggota", "Digit ke-6", "Segmen", "Nominal Tunggakan (Rp)"],
    ];

    filteredPelanggan.forEach((p, idx) => {
      semuaData.push([
        idx + 1,
        p.idpel,
        p.nama,
        p.tarifDaya,
        p.rbm,
        p.regu,
        (p as any).kodeAnggota || p.idPetugas,
        (p as any).kodePerorangan || "-",
        p.kategori,
        p.rpTunggakan,
      ]);
    });

    const ws3 = XLSX.utils.aoa_to_sheet(semuaData);
    ws3["!cols"] = [{ wch: 6 }, { wch: 18 }, { wch: 35 }, { wch: 16 }, { wch: 16 }, { wch: 10 }, { wch: 14 }, { wch: 10 }, { wch: 12 }, { wch: 24 }];
    XLSX.utils.book_append_sheet(wb, ws3, "Data Lengkap Pelanggan");

    XLSX.writeFile(
      wb,
      `PLN_UP3_Jatinegara_KOGOL_${activeData.periode.replace(/\s+/g, "_")}_Regu_${filterRegu}.xlsx`
    );
  };

  const donutPelanggan = activeData
    ? [
        { name: "AMR", value: activeData.denganKcic?.amrPlgn || 0, color: "#005C8A" },
        { name: "NON-AMR", value: activeData.denganKcic?.nonAmrPlgn || 0, color: "#E2E8F0" },
      ]
    : [];

  const donutRupiah = activeData
    ? [
        { name: "AMR", value: activeData.denganKcic?.amrRp || 0, color: "#005C8A" },
        { name: "NON-AMR", value: activeData.denganKcic?.nonAmrRp || 0, color: "#EA580C" },
      ]
    : [];

  return (
    <main className="min-h-screen bg-[#F4F6F9] text-slate-800 font-sans pb-16 antialiased overflow-x-hidden print:bg-white print:p-2">
      {/* KOP CETAK RESMI PLN */}
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
              <div>No. Dokumen: <strong>PLN-UP3JTN/KOGOL/{activeData.tahun}/{activeData.kodeBulan}</strong></div>
              <div>Penugasan: <strong>Regu {filterRegu} / Petugas {filterPetugas}</strong></div>
              <div>Tanggal Cetak: <strong>{waktuSinkronisasi}</strong></div>
            </div>
          </div>
          <div className="text-center mt-3 pt-2 border-t border-slate-200">
            <h1 className="text-sm font-black text-slate-900 uppercase">
              LEMBAR PENUGASAN PENAGIHAN LAPANGAN (KOGOL 0)
            </h1>
            <p className="text-xs font-semibold text-slate-700">Periode: {activeData.periode.toUpperCase()}</p>
          </div>
        </div>
      )}

      {/* TOP NAVBAR ENTERPRISE BUMN (RESPONSIF MOBILE & DESKTOP) */}
      <header className="sticky top-0 z-40 bg-white border-b border-slate-200/90 shadow-2xs print:hidden">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-2.5 sm:py-0 sm:h-16 flex flex-col sm:flex-row sm:items-center justify-between gap-2.5">
          <div className="flex items-center gap-2.5 min-w-0">
            <PlnLogo className="w-7 h-9 shrink-0" />
            <div className="min-w-0">
              <div className="flex items-center gap-1.5 text-[10px] sm:text-[11px] truncate">
                <span className="font-black tracking-widest text-[#005C8A] uppercase">PT PLN (Persero)</span>
                <span className="text-slate-300">/</span>
                <span className="font-semibold text-slate-500 truncate">UP3 Jatinegara</span>
              </div>
              <h1 className="text-xs sm:text-sm font-bold text-slate-900 tracking-tight truncate">
                SATU JATINEGARA — KOGOL Suite
              </h1>
            </div>
          </div>

          <div className="flex items-center gap-1.5 overflow-x-auto pb-1 sm:pb-0 scrollbar-none w-full sm:w-auto">
            {activeData && (
              <>
                <button
                  onClick={() => window.print()}
                  className="shrink-0 inline-flex items-center gap-1 px-2.5 py-1.5 text-[11px] sm:text-xs font-semibold text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 rounded-lg shadow-2xs transition"
                >
                  <Printer className="w-3.5 h-3.5 text-slate-500" />
                  <span className="hidden sm:inline">Cetak (PDF)</span>
                  <span className="sm:hidden">PDF</span>
                </button>
                <button
                  onClick={handleDownloadExcel}
                  className="shrink-0 inline-flex items-center gap-1 px-2.5 py-1.5 text-[11px] sm:text-xs font-semibold text-white bg-emerald-600 hover:bg-emerald-700 rounded-lg shadow-2xs transition"
                >
                  <FileSpreadsheet className="w-3.5 h-3.5" />
                  <span>Excel</span>
                </button>
                <button
                  onClick={() => downloadPresentationPptx(activeData, modeLaporan, dataHistory)}
                  className="shrink-0 inline-flex items-center gap-1 px-2.5 py-1.5 text-[11px] sm:text-xs font-semibold text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 rounded-lg shadow-2xs transition"
                >
                  <Presentation className="w-3.5 h-3.5 text-amber-600" />
                  <span>PPTX</span>
                </button>
              </>
            )}

            <label className="shrink-0 cursor-pointer inline-flex items-center gap-1 px-3 py-1.5 text-[11px] sm:text-xs font-bold text-white bg-[#005C8A] hover:bg-[#00476B] rounded-lg shadow-2xs transition">
              {sortedKeys.length === 0 ? (
                <>
                  <Upload className="w-3.5 h-3.5" />
                  <span>Unggah</span>
                </>
              ) : (
                <>
                  <PlusCircle className="w-3.5 h-3.5" />
                  <span>+ Bulan</span>
                </>
              )}
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls"
                multiple
                className="hidden"
                onChange={handleFileUpload}
              />
            </label>

            {sortedKeys.length > 0 && (
              <button
                onClick={handleResetData}
                title="Reset Semua Data"
                className="shrink-0 p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg border border-slate-200 transition"
              >
                <RotateCcw className="w-3.5 h-3.5" />
              </button>
            )}
          </div>
        </div>
      </header>

      <div className="max-w-7xl mx-auto px-3 sm:px-6 lg:px-8 pt-4 sm:pt-6 space-y-4 sm:space-y-6">
        {/* SUB-BAR OPERASIONAL STATUS */}
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 bg-white px-3.5 py-2 rounded-xl border border-slate-200/80 shadow-2xs text-[11px] sm:text-xs print:hidden">
          <div className="flex items-center gap-2 flex-wrap">
            <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded-md bg-emerald-50 text-emerald-700 font-bold border border-emerald-200">
              <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse" />
              KOGOL 0 AKHIR
            </span>
            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md bg-sky-50 text-[#005C8A] font-semibold border border-sky-200">
              <ShieldCheck className="w-3.5 h-3.5" />
              5 IDPEL KCIC Terverifikasi
            </span>
            {activeData && (
              <span className="text-slate-500">
                Periode: <strong className="text-slate-900">{activeData.periode}</strong>
              </span>
            )}
          </div>
          <div className="text-slate-400 flex items-center gap-1 text-[10px] sm:text-xs">
            <Clock className="w-3 h-3" />
            <span>Sinkronisasi: {waktuSinkronisasi || "-"}</span>
          </div>
        </div>

        {sortedKeys.length === 0 ? (
          <div className="bg-white border-2 border-dashed border-slate-300 rounded-2xl p-10 sm:p-16 text-center flex flex-col items-center justify-center space-y-3 shadow-2xs">
            <div className="w-12 h-12 bg-sky-50 text-[#005C8A] rounded-xl flex items-center justify-center border border-sky-100">
              <Upload className="w-6 h-6" />
            </div>
            <h3 className="text-sm font-bold text-slate-800">Menunggu Berkas Excel KOGOL</h3>
            <p className="text-xs text-slate-500 max-w-sm">
              Klik tombol &quot;Unggah&quot; untuk memilih satu atau beberapa berkas Excel resmi KOGOL sekaligus.
            </p>
          </div>
        ) : (
          activeData && (
            <>
              {/* TAB SELECTOR PERIODE & TOGGLE DASHBOARD/LEMBAR KERJA */}
              <div className="flex flex-col md:flex-row justify-between items-stretch md:items-center gap-2.5 bg-white p-2 sm:p-2.5 rounded-xl border border-slate-200 shadow-2xs print:hidden">
                <div className="flex items-center gap-1.5 overflow-x-auto pb-1 md:pb-0 scrollbar-none">
                  <span className="text-[10px] sm:text-[11px] font-bold text-slate-400 uppercase tracking-wider pr-1">Bulan:</span>
                  {sortedKeys.map((k) => (
                    <button
                      key={k}
                      onClick={() => {
                        setSelectedKey(k);
                        setSimulasiLunas({});
                      }}
                      className={`shrink-0 px-2.5 sm:px-3 py-1 rounded-lg text-[11px] sm:text-xs font-bold transition ${
                        k === selectedKey
                          ? "bg-[#005C8A] text-white shadow-2xs"
                          : "text-slate-600 hover:bg-slate-100"
                      }`}
                    >
                      {dataHistory[k]?.periode || k}
                    </button>
                  ))}
                </div>

                <div className="flex bg-slate-100 p-1 rounded-lg border border-slate-200 text-xs font-semibold shrink-0">
                  <button
                    onClick={() => setActiveTab("RINGKASAN")}
                    className={`flex-1 sm:flex-initial flex items-center justify-center gap-1.5 px-3 py-1 rounded-md transition ${
                      activeTab === "RINGKASAN" ? "bg-white text-[#005C8A] shadow-xs font-bold" : "text-slate-500 hover:text-slate-700"
                    }`}
                  >
                    <Building2 className="w-3.5 h-3.5" />
                    <span>Eksekutif</span>
                  </button>
                  <button
                    onClick={() => setActiveTab("LEMBAR_KERJA")}
                    className={`flex-1 sm:flex-initial flex items-center justify-center gap-1.5 px-3 py-1 rounded-md transition ${
                      activeTab === "LEMBAR_KERJA" ? "bg-white text-[#005C8A] shadow-xs font-bold" : "text-slate-500 hover:text-slate-700"
                    }`}
                  >
                    <FileText className="w-3.5 h-3.5" />
                    <span>Lembar Kerja</span>
                  </button>
                </div>
              </div>

              {activeTab === "RINGKASAN" && (
                <>
                  {/* BAR FILTER SKENARIO WHAT-IF */}
                  <div className="bg-white p-3 sm:p-3.5 rounded-xl border border-slate-200 shadow-2xs space-y-2.5 print:hidden">
                    <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-2.5">
                      <div className="flex items-center gap-2 text-xs">
                        <div className="p-1 rounded-lg bg-sky-50 text-[#005C8A]">
                          <Filter className="w-3.5 h-3.5" />
                        </div>
                        <span className="font-bold text-slate-800 uppercase tracking-wider text-[11px]">Skenario Portofolio:</span>
                      </div>

                      <div className="flex bg-slate-100 p-1 rounded-lg border border-slate-200 text-xs font-semibold overflow-x-auto w-full lg:w-auto scrollbar-none">
                        <button
                          onClick={() => setSkenario("SEMUA")}
                          className={`shrink-0 flex-1 sm:flex-initial px-3 py-1 rounded-md text-[11px] sm:text-xs transition ${
                            skenario === "SEMUA" ? "bg-white text-slate-900 shadow-2xs font-bold" : "text-slate-500 hover:text-slate-700"
                          }`}
                        >
                          Semua
                        </button>
                        <button
                          onClick={() => setSkenario("TANPA_KCIC")}
                          className={`shrink-0 flex-1 sm:flex-initial px-3 py-1 rounded-md text-[11px] sm:text-xs transition ${
                            skenario === "TANPA_KCIC" ? "bg-white text-[#005C8A] shadow-2xs font-bold" : "text-slate-500 hover:text-slate-700"
                          }`}
                        >
                          Tanpa KCIC (5 IDPEL)
                        </button>
                        <button
                          onClick={() => setSkenario("HANYA_NON_AMR")}
                          className={`shrink-0 flex-1 sm:flex-initial px-3 py-1 rounded-md text-[11px] sm:text-xs transition ${
                            skenario === "HANYA_NON_AMR" ? "bg-white text-emerald-700 shadow-2xs font-bold" : "text-slate-500 hover:text-slate-700"
                          }`}
                        >
                          Non-AMR
                        </button>
                      </div>
                    </div>

                    {totalLunasSimulasi > 0 && (
                      <div className="p-2.5 bg-emerald-50 border border-emerald-200 rounded-lg flex flex-col sm:flex-row justify-between items-start sm:items-center gap-2 text-xs text-emerald-900">
                        <span className="flex items-center gap-2">
                          <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
                          <span>
                            Simulasi: <b>{jumlahPelangganLunasSimulasi} Pelanggan</b> lunas (Pengurangan: <b>Rp {totalLunasSimulasi.toLocaleString("id-ID")}</b>).
                          </span>
                        </span>
                        <button
                          onClick={() => setSimulasiLunas({})}
                          className="text-[11px] underline font-bold hover:text-emerald-700 self-end sm:self-auto"
                        >
                          Reset Simulasi
                        </button>
                      </div>
                    )}
                  </div>

                  {/* BENTO GRID ROW 1: BANNER TARGET POLARITAS NEGATIF & DUAL METRIC CARDS */}
                  <div className="grid grid-cols-1 lg:grid-cols-12 gap-3.5 sm:gap-5">
                    <div className="lg:col-span-4 bg-[#003B5C] rounded-2xl p-4 sm:p-5 text-white shadow-xs border border-sky-950 flex flex-col justify-between">
                      <div>
                        <div className="flex items-center justify-between">
                          <span className="text-[10px] sm:text-[11px] font-bold text-sky-200 uppercase tracking-wider flex items-center gap-1.5">
                            <Target className="w-3.5 h-3.5 text-amber-400" />
                            Pencapaian Target Saldo
                          </span>
                          <span className="text-[9px] sm:text-[10px] px-2 py-0.5 rounded bg-sky-900/80 text-sky-200 border border-sky-700/50 font-bold">
                            {skenario.replace(/_/g, " ")}
                          </span>
                        </div>
                        <div className="text-3xl sm:text-4xl font-black text-rose-400 mt-2.5 sm:mt-3 tracking-tight">
                          {capaianPersenSimulasi}%
                        </div>
                        <p className="text-[10px] sm:text-[11px] text-sky-200/70 mt-1">
                          Polaritas Negatif: 2 - (Realisasi : Target)
                        </p>
                      </div>

                      <div className="mt-4 sm:mt-5 pt-3 border-t border-sky-800/80 grid grid-cols-2 gap-2 text-xs">
                        <div>
                          <span className="text-sky-300/80 text-[10px] block uppercase font-bold">Batas Target Unit</span>
                          <strong className="text-white">Rp {targetSaldo.toLocaleString("id-ID")}</strong>
                        </div>
                        <div className="text-right">
                          <span className="text-sky-300/80 text-[10px] block uppercase font-bold">Sisa Tunggakan</span>
                          <strong className="text-amber-300">Rp {saldoSimulasi.toLocaleString("id-ID")}</strong>
                        </div>
                      </div>
                    </div>

                    <div className="lg:col-span-4 bg-white rounded-2xl p-4 sm:p-5 border border-slate-200 shadow-2xs flex flex-col justify-between">
                      <div>
                        <div className="flex items-center justify-between border-b border-slate-100 pb-2.5 mb-2.5">
                          <div className="flex items-center gap-2">
                            <div className="w-2 h-2 rounded-full bg-[#005C8A]" />
                            <h3 className="text-xs font-bold text-slate-800 uppercase tracking-wider">Total Tagihan (Dengan KCIC)</h3>
                          </div>
                          <span className="text-[10px] font-bold px-2 py-0.5 rounded bg-sky-50 text-[#005C8A] border border-sky-100">
                            Resmi
                          </span>
                        </div>

                        <span className="text-[10px] sm:text-[11px] text-slate-400 font-medium block uppercase">Total Piutang Portofolio</span>
                        <div className="text-xl sm:text-2xl font-black text-[#005C8A] mt-0.5">
                          Rp {(activeData.denganKcic?.totalRp || 0).toLocaleString("id-ID")}
                        </div>
                        <span className="text-xs text-slate-500 mt-1 block">
                          Beban Pelanggan: <strong className="text-slate-800">{activeData.denganKcic?.totalPlgn || 0} IDPEL</strong>
                        </span>
                      </div>

                      <div className="mt-3.5 pt-2.5 border-t border-slate-100 grid grid-cols-2 text-xs text-slate-600 bg-slate-50/70 p-2 rounded-xl">
                        <div>AMR: <strong>Rp {(activeData.denganKcic?.amrRp || 0).toLocaleString("id-ID")}</strong></div>
                        <div className="text-right">Non-AMR: <strong>Rp {(activeData.denganKcic?.nonAmrRp || 0).toLocaleString("id-ID")}</strong></div>
                      </div>
                    </div>

                    <div className="lg:col-span-4 bg-white rounded-2xl p-4 sm:p-5 border border-slate-200 shadow-2xs flex flex-col justify-between">
                      <div>
                        <div className="flex items-center justify-between border-b border-slate-100 pb-2.5 mb-2.5">
                          <div className="flex items-center gap-2">
                            <div className="w-2 h-2 rounded-full bg-amber-500" />
                            <h3 className="text-xs font-bold text-slate-800 uppercase tracking-wider">Total Tagihan (Tanpa KCIC)</h3>
                          </div>
                          <span className="text-[10px] font-bold px-2 py-0.5 rounded bg-amber-50 text-amber-700 border border-amber-100">
                            Reguler
                          </span>
                        </div>

                        <span className="text-[10px] sm:text-[11px] text-slate-400 font-medium block uppercase">Piutang Murni Reguler</span>
                        <div className="text-xl sm:text-2xl font-black text-slate-900 mt-0.5">
                          Rp {(activeData.tanpaKcic?.totalRp || 0).toLocaleString("id-ID")}
                        </div>
                        <span className="text-xs text-slate-500 mt-1 block">
                          Beban Pelanggan: <strong className="text-slate-800">{activeData.tanpaKcic?.totalPlgn || 0} IDPEL</strong>
                        </span>
                      </div>

                      <div className="mt-3.5 pt-2.5 border-t border-slate-100 grid grid-cols-2 text-xs text-slate-600 bg-slate-50/70 p-2 rounded-xl">
                        <div>AMR: <strong>Rp {(activeData.tanpaKcic?.amrRp || 0).toLocaleString("id-ID")}</strong></div>
                        <div className="text-right">Non-AMR: <strong>Rp {(activeData.tanpaKcic?.nonAmrRp || 0).toLocaleString("id-ID")}</strong></div>
                      </div>
                    </div>
                  </div>

                  {/* BENTO GRID ROW 2: SEGMENTASI GOLONGAN TARIF */}
                  {activeData.rekapTarif && (
                    <div className="bg-white rounded-2xl p-4 sm:p-5 border border-slate-200 shadow-2xs">
                      <div className="flex items-center justify-between mb-3 border-b border-slate-100 pb-2">
                        <h3 className="text-xs font-bold text-slate-800 uppercase tracking-wider flex items-center gap-1.5">
                          <Layers className="w-4 h-4 text-[#005C8A]" />
                          Segmentasi Tarif ({activeData.periode})
                        </h3>
                        <span className="text-[10px] sm:text-[11px] text-slate-400">Rincian Kelompok Tarif</span>
                      </div>
                      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-2.5 sm:gap-3">
                        {[
                          { label: "Rumah Tangga (R)", key: "R", badge: "bg-sky-50 text-sky-700 border-sky-200" },
                          { label: "Bisnis (B)", key: "B", badge: "bg-amber-50 text-amber-700 border-amber-200" },
                          { label: "Industri (I)", key: "I", badge: "bg-purple-50 text-purple-700 border-purple-200" },
                          { label: "Publik / Kantor (P)", key: "P", badge: "bg-emerald-50 text-emerald-700 border-emerald-200" },
                          { label: "Kereta Cepat (T / KCIC)", key: "T", badge: "bg-rose-50 text-rose-700 border-rose-200" },
                        ].map((t) => {
                          const val = activeData.rekapTarif[t.key] || { plgn: 0, rp: 0 };
                          return (
                            <div key={t.key} className="p-2.5 sm:p-3 rounded-xl border border-slate-100 bg-slate-50/60">
                              <span className={`text-[9px] sm:text-[10px] font-bold px-1.5 sm:px-2 py-0.5 rounded-md inline-block border ${t.badge}`}>
                                {t.label}
                              </span>
                              <div className="text-sm sm:text-base font-black text-slate-900 mt-1.5">
                                Rp {val.rp.toLocaleString("id-ID")}
                              </div>
                              <span className="text-[10px] sm:text-[11px] text-slate-500">{val.plgn} Pelanggan</span>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  )}

                  {/* BENTO GRID ROW 3: TABEL PEMANTAUAN BULANAN & LEADERBOARD REGU */}
                  <div className="grid grid-cols-1 lg:grid-cols-12 gap-3.5 sm:gap-5">
                    <div className="lg:col-span-7 bg-white rounded-2xl border border-slate-200 shadow-2xs overflow-hidden flex flex-col justify-between">
                      <div className="px-4 sm:px-5 py-3 sm:py-4 border-b border-slate-100 bg-slate-50/50 flex items-center justify-between">
                        <div>
                          <h3 className="text-xs font-bold text-slate-800 uppercase tracking-wider">
                            Perbandingan & Rata-Rata Tagihan Antar Bulan
                          </h3>
                          <p className="text-[10px] sm:text-[11px] text-slate-400 mt-0.5">
                            Polaritas Negatif: <b>2 - (Realisasi Dengan KCIC : Target)</b>
                          </p>
                        </div>
                      </div>

                      <div className="overflow-x-auto">
                        <table className="w-full text-xs text-left">
                          <thead className="bg-slate-50/80 text-slate-600 font-bold border-b border-slate-100">
                            <tr>
                              <th className="py-2.5 px-3 sm:px-4">Periode</th>
                              <th className="py-2.5 px-2 sm:px-3 text-center">Plgn</th>
                              <th className="py-2.5 px-3 sm:px-4 text-right">Total (Dengan KCIC)</th>
                              <th className="py-2.5 px-3 sm:px-4 text-right">Capaian (%)</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100 text-slate-700">
                            {perbandinganBulanan.list.map((b) => (
                              <tr key={b.key} className={b.key === selectedKey ? "bg-sky-50/40 font-medium" : "hover:bg-slate-50/50"}>
                                <td className="py-2.5 px-3 sm:px-4 text-slate-900 font-bold whitespace-nowrap">
                                  {b.periode}
                                  <span className="ml-1 text-[10px] text-slate-400 font-normal">(Bln {b.pembagiBulan})</span>
                                </td>
                                <td className="py-2.5 px-2 sm:px-3 text-center">{b.totalPlgn}</td>
                                <td className="py-2.5 px-3 sm:px-4 text-right font-bold text-[#005C8A] whitespace-nowrap">
                                  Rp {b.totalRp.toLocaleString("id-ID")}
                                </td>
                                <td className="py-2.5 px-3 sm:px-4 text-right">
                                  <span
                                    className={`inline-block px-1.5 sm:px-2 py-0.5 rounded-md font-black text-[11px] sm:text-xs ${
                                      b.persen >= 100
                                        ? "bg-emerald-50 text-emerald-700 border border-emerald-200"
                                        : "bg-rose-50 text-rose-700 border border-rose-200"
                                    }`}
                                  >
                                    {b.persen.toFixed(1)}%
                                  </span>
                                </td>
                              </tr>
                            ))}
                          </tbody>
                          <tfoot className="bg-slate-50 font-black text-slate-900 border-t border-slate-200">
                            <tr>
                              <td className="py-2.5 px-3 sm:px-4">TOTAL ({perbandinganBulanan.jumlahBulan} Bln)</td>
                              <td className="py-2.5 px-2 sm:px-3 text-center">{perbandinganBulanan.totalPlgnGabungan}</td>
                              <td className="py-2.5 px-3 sm:px-4 text-right text-[#005C8A] whitespace-nowrap">
                                Rp {perbandinganBulanan.totalRpGabungan.toLocaleString("id-ID")}
                              </td>
                              <td className="py-2.5 px-3 sm:px-4 text-right text-slate-400">—</td>
                            </tr>
                            <tr className="bg-sky-50/60 text-[#005C8A] font-bold border-t border-slate-100">
                              <td className="py-2.5 px-3 sm:px-4">RATA-RATA / BLN</td>
                              <td className="py-2.5 px-2 sm:px-3 text-center">{perbandinganBulanan.rataRataPlgn}</td>
                              <td className="py-2.5 px-3 sm:px-4 text-right whitespace-nowrap">
                                Rp {Math.round(perbandinganBulanan.rataRataRp).toLocaleString("id-ID")}
                              </td>
                              <td className="py-2.5 px-3 sm:px-4 text-right font-black">
                                {perbandinganBulanan.persenTotalAkhir.toFixed(1)}%
                              </td>
                            </tr>
                          </tfoot>
                        </table>
                      </div>
                    </div>

                    {/* Peringkat Beban Per Regu (Menempel rapi ke atas) */}
                    <div className="lg:col-span-5 bg-white rounded-2xl border border-slate-200 shadow-2xs overflow-hidden flex flex-col justify-start">
                      <div className="px-4 sm:px-5 py-3 sm:py-4 border-b border-slate-100 bg-slate-50/50 flex items-center justify-between">
                        <div className="flex items-center gap-1.5">
                          <Trophy className="w-4 h-4 text-amber-500" />
                          <h3 className="text-xs font-bold text-slate-800 uppercase tracking-wider">
                            Peringkat Regu ({activeData.periode})
                          </h3>
                        </div>
                        <span className="text-[10px] sm:text-[11px] text-slate-400">Tertinggi ke Terendah</span>
                      </div>

                      <div className="overflow-x-auto flex-1">
                        <table className="w-full text-xs text-left">
                          <thead className="bg-slate-50/80 text-slate-600 font-bold border-b border-slate-100">
                            <tr>
                              <th className="py-2.5 px-2 text-center w-8">No</th>
                              <th className="py-2.5 px-3">Regu</th>
                              <th className="py-2.5 px-2 text-center">Plgn</th>
                              <th className="py-2.5 px-3 sm:px-4 text-right">Tunggakan (Rp)</th>
                              <th className="py-2.5 px-2.5 text-right">Porsi</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100 text-slate-700">
                            {peringkatRegu.list.map((r, idx) => {
                              const porsi = ((r.totalRp / (peringkatRegu.totalRp || 1)) * 100).toFixed(1);
                              return (
                                <tr key={r.regu} className="hover:bg-slate-50/50 transition">
                                  <td className="py-2 px-2 text-center font-bold text-slate-400">{idx + 1}</td>
                                  <td className="py-2 px-3 font-bold text-[#005C8A]">REGU {r.regu}</td>
                                  <td className="py-2 px-2 text-center">{r.totalPlgn}</td>
                                  <td className="py-2 px-3 sm:px-4 text-right font-black text-slate-900 whitespace-nowrap">
                                    Rp {r.totalRp.toLocaleString("id-ID")}
                                  </td>
                                  <td className="py-2 px-2.5 text-right font-semibold text-slate-500">{porsi}%</td>
                                </tr>
                              );
                            })}
                          </tbody>
                          <tfoot className="bg-slate-50 font-black text-slate-900 border-t border-slate-200">
                            <tr>
                              <td colSpan={2} className="py-2.5 px-3 uppercase">Total</td>
                              <td className="py-2.5 px-2 text-center">{peringkatRegu.totalPlgn}</td>
                              <td className="py-2.5 px-3 sm:px-4 text-right text-[#005C8A] whitespace-nowrap">
                                Rp {peringkatRegu.totalRp.toLocaleString("id-ID")}
                              </td>
                              <td className="py-2.5 px-2.5 text-right">100%</td>
                            </tr>
                          </tfoot>
                        </table>
                      </div>
                    </div>
                  </div>

                  {/* BENTO GRID ROW 4: ANALISIS KOMPARASI KINERJA (MoM) */}
                  {hasComparison && prevData && currData && (
                    <div className="bg-white rounded-2xl border border-slate-200 shadow-2xs overflow-hidden">
                      <div className="bg-gradient-to-r from-slate-900 to-[#003B5C] px-4 sm:px-6 py-3.5 sm:py-4 text-white flex flex-col sm:flex-row justify-between sm:items-center gap-2">
                        <div>
                          <h3 className="text-xs sm:text-sm font-black tracking-wider uppercase flex items-center gap-1.5">
                            <span>📊</span> Komparasi: {prevData.periode} vs {currData.periode}
                          </h3>
                          <p className="text-[11px] text-sky-200">
                            Evaluasi perubahan saldo piutang tunggakan antar-bulan secara otomatis
                          </p>
                        </div>
                        <div className="flex items-center gap-2 text-xs bg-white/10 px-2.5 py-1 rounded-lg backdrop-blur-xs font-semibold self-start sm:self-auto">
                          <span>{(prevData.periode || "").split(" ")[0]}</span>
                          <ArrowRight className="w-3.5 h-3.5" />
                          <span className="text-[#FFD100]">{(currData.periode || "").split(" ")[0]}</span>
                        </div>
                      </div>

                      <div className="p-3.5 sm:p-5 grid grid-cols-1 sm:grid-cols-3 gap-3 sm:gap-4 border-b border-slate-100 bg-slate-50/50">
                        <div className="bg-white p-3 sm:p-4 rounded-xl border border-slate-200 shadow-2xs">
                          <span className="text-[10px] sm:text-[11px] font-bold text-slate-500 uppercase tracking-wider">
                            Selisih Total (Dengan KCIC)
                          </span>
                          <div className="flex items-baseline gap-2 mt-1.5">
                            <span className={`text-lg sm:text-xl font-black ${diffTotalRp > 0 ? "text-rose-600" : "text-emerald-600"}`}>
                              {diffTotalRp > 0 ? "+" : ""}Rp {Math.abs(diffTotalRp).toLocaleString("id-ID")}
                            </span>
                          </div>
                          <div className="flex items-center gap-1 mt-1.5">
                            {diffTotalRp > 0 ? (
                              <span className="inline-flex items-center gap-1 text-[10px] sm:text-[11px] font-bold px-1.5 py-0.5 rounded bg-rose-50 text-rose-600">
                                <TrendingUp className="w-3 h-3" /> +{pctTotalRp.toFixed(1)}% (Naik)
                              </span>
                            ) : (
                              <span className="inline-flex items-center gap-1 text-[10px] sm:text-[11px] font-bold px-1.5 py-0.5 rounded bg-emerald-50 text-emerald-600">
                                <TrendingDown className="w-3 h-3" /> {pctTotalRp.toFixed(1)}% (Turun)
                              </span>
                            )}
                          </div>
                        </div>

                        <div className="bg-white p-3 sm:p-4 rounded-xl border border-slate-200 shadow-2xs">
                          <span className="text-[10px] sm:text-[11px] font-bold text-slate-500 uppercase tracking-wider">
                            Selisih Murni (Tanpa KCIC)
                          </span>
                          <div className="flex items-baseline gap-2 mt-1.5">
                            <span className={`text-lg sm:text-xl font-black ${diffTanpaKcicRp > 0 ? "text-rose-600" : "text-emerald-600"}`}>
                              {diffTanpaKcicRp > 0 ? "+" : ""}Rp {Math.abs(diffTanpaKcicRp).toLocaleString("id-ID")}
                            </span>
                          </div>
                          <div className="flex items-center gap-1 mt-1.5">
                            {diffTanpaKcicRp > 0 ? (
                              <span className="inline-flex items-center gap-1 text-[10px] sm:text-[11px] font-bold px-1.5 py-0.5 rounded bg-rose-50 text-rose-600">
                                <TrendingUp className="w-3 h-3" /> +{pctTanpaKcicRp.toFixed(1)}% (Naik)
                              </span>
                            ) : (
                              <span className="inline-flex items-center gap-1 text-[10px] sm:text-[11px] font-bold px-1.5 py-0.5 rounded bg-emerald-50 text-emerald-600">
                                <TrendingDown className="w-3 h-3" /> {pctTanpaKcicRp.toFixed(1)}% (Turun)
                              </span>
                            )}
                          </div>
                        </div>

                        <div className="bg-white p-3 sm:p-4 rounded-xl border border-slate-200 shadow-2xs">
                          <span className="text-[10px] sm:text-[11px] font-bold text-slate-500 uppercase tracking-wider">
                            Perubahan Jumlah Pelanggan
                          </span>
                          <div className="flex items-baseline gap-2 mt-1.5">
                            <span className={`text-lg sm:text-xl font-black ${diffPlgn > 0 ? "text-amber-600" : "text-emerald-600"}`}>
                              {diffPlgn > 0 ? `+${diffPlgn}` : diffPlgn} Plgn
                            </span>
                          </div>
                          <div className="text-[11px] text-slate-500 mt-1.5">
                            Dari <b>{prevData.denganKcic?.totalPlgn || 0}</b> ke <b>{currData.denganKcic?.totalPlgn || 0}</b>
                          </div>
                        </div>
                      </div>

                      <div className="p-3.5 sm:p-5">
                        <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5 sm:gap-3">
                          <div className="bg-slate-50 border border-slate-200 rounded-xl p-3 space-y-1">
                            <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider block">
                              🚆 Portofolio KCIC (5 IDPEL)
                            </span>
                            <div className="text-xs text-slate-800 leading-relaxed">
                              {diffKcicRp > 0 ? (
                                <>
                                  Saldo bertambah <b>Rp {diffKcicRp.toLocaleString("id-ID")}</b> ({kcicContributionPct}% dari kenaikan total).
                                </>
                              ) : diffKcicRp < 0 ? (
                                <>
                                  Saldo turun <b>Rp {Math.abs(diffKcicRp).toLocaleString("id-ID")}</b>.
                                </>
                              ) : (
                                "Saldo terpantau stagnan."
                              )}
                            </div>
                          </div>

                          <div className="bg-slate-50 border border-slate-200 rounded-xl p-3 space-y-1">
                            <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider block">
                              ⚡ Saldo Murni Reguler
                            </span>
                            <div className="text-xs text-slate-800 leading-relaxed">
                              {diffTanpaKcicRp > 0 ? (
                                <>
                                  Mengalami kenaikan sebesar <b className="text-rose-600">Rp {diffTanpaKcicRp.toLocaleString("id-ID")}</b> (+{pctTanpaKcicRp.toFixed(1)}%).
                                </>
                              ) : (
                                <>
                                  Terkendali dengan reduksi <b className="text-emerald-600">Rp {Math.abs(diffTanpaKcicRp).toLocaleString("id-ID")}</b> ({pctTanpaKcicRp.toFixed(1)}%).
                                </>
                              )}
                            </div>
                          </div>

                          <div className="bg-slate-50 border border-slate-200 rounded-xl p-3 space-y-1">
                            <span className="text-[10px] font-bold text-[#005C8A] uppercase tracking-wider block">
                              📋 Rekomendasi
                            </span>
                            <div className="text-xs text-slate-700 leading-relaxed">
                              {diffTotalRp > 0
                                ? "Prioritaskan rekonsiliasi faktur 5 IDPEL KCIC terpusat dan lakukan tindakan penagihan untuk pelanggan reguler berpiutang besar."
                                : "Pertahankan ritme penagihan harian dan pantau pengawasan penyegelan pelanggan yang belum melunasi kewajibannya."}
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>
                  )}

                  {/* BENTO GRID ROW 5: VISUALISASI GRAFIK */}
                  <div className="grid grid-cols-1 lg:grid-cols-12 gap-3.5 sm:gap-5 print:hidden">
                    <div className="lg:col-span-7 bg-white rounded-2xl p-4 sm:p-5 border border-slate-200 shadow-2xs">
                      <div className="flex items-center justify-between mb-3 border-b border-slate-100 pb-2">
                        <h3 className="text-xs font-bold text-slate-800 uppercase tracking-wider flex items-center gap-1.5">
                          <BarChart3 className="w-4 h-4 text-[#005C8A]" />
                          Trend Realisasi Multi-Bulan
                        </h3>
                      </div>
                      <div className="h-52 sm:h-56">
                        <ResponsiveContainer width="100%" height="100%">
                          <BarChart data={barChartData}>
                            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#F1F5F9" />
                            <XAxis dataKey="bulan" stroke="#94A3B8" fontSize={11} />
                            <YAxis tickFormatter={(val) => `${(val / 1e9).toFixed(1)}M`} stroke="#94A3B8" fontSize={11} />
                            <Tooltip formatter={(val: any) => `Rp ${Number(val).toLocaleString("id-ID")}`} />
                            <Bar dataKey="Dengan_KCIC" fill="#005C8A" radius={[4, 4, 0, 0]} name="Dengan KCIC" maxBarSize={32} />
                            <Bar dataKey="Tanpa_KCIC" fill="#F59E0B" radius={[4, 4, 0, 0]} name="Tanpa KCIC" maxBarSize={32} />
                          </BarChart>
                        </ResponsiveContainer>
                      </div>
                    </div>

                    <div className="lg:col-span-5 grid grid-cols-2 gap-2.5 sm:gap-4">
                      <div className="bg-white rounded-2xl p-3 sm:p-4 border border-slate-200 shadow-2xs flex flex-col justify-between">
                        <span className="text-[10px] sm:text-[11px] font-bold text-slate-700 uppercase block border-b border-slate-100 pb-1.5 truncate">
                          % Pelanggan AMR
                        </span>
                        <div className="h-36 sm:h-44">
                          <ResponsiveContainer width="100%" height="100%">
                            <PieChart>
                              <Pie data={donutPelanggan} dataKey="value" innerRadius={34} outerRadius={52} paddingAngle={3}>
                                {donutPelanggan.map((entry, idx) => (
                                  <Cell key={`cell-${idx}`} fill={entry.color} />
                                ))}
                              </Pie>
                              <Tooltip />
                            </PieChart>
                          </ResponsiveContainer>
                        </div>
                      </div>

                      <div className="bg-white rounded-2xl p-3 sm:p-4 border border-slate-200 shadow-2xs flex flex-col justify-between">
                        <span className="text-[10px] sm:text-[11px] font-bold text-slate-700 uppercase block border-b border-slate-100 pb-1.5 truncate">
                          % Rupiah AMR
                        </span>
                        <div className="h-36 sm:h-44">
                          <ResponsiveContainer width="100%" height="100%">
                            <PieChart>
                              <Pie data={donutRupiah} dataKey="value" innerRadius={34} outerRadius={52} paddingAngle={3}>
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
                  </div>
                </>
              )}

              {/* BENTO GRID ROW 6: TABEL RINCIAN PELANGGAN & LEMBAR KERJA */}
              <div className="bg-white rounded-2xl border border-slate-200 shadow-2xs overflow-hidden print:border-none">
                <div className="px-4 sm:px-5 py-3 sm:py-4 border-b border-slate-100 flex flex-col gap-2.5 bg-slate-50/50 print:bg-white print:p-0 print:border-b">
                  <div>
                    <h3 className="text-xs font-bold text-slate-800 uppercase tracking-wider flex items-center gap-1.5">
                      <Users className="w-4 h-4 text-[#005C8A]" />
                      {activeTab === "LEMBAR_KERJA"
                        ? `Lembar Penugasan Lapangan — ${activeData.periode}`
                        : `Daftar Tunggakan Pelanggan & Alokasi RBM — ${activeData.periode}`}
                    </h3>
                    <p className="text-[10px] sm:text-[11px] text-slate-400 mt-0.5 print:text-slate-600">
                      Format RBM: HAAMRTI... ➔ Regu: <b>MR</b>, Kode Anggota: <b>MRT</b>
                    </p>
                  </div>

                  <div className="grid grid-cols-2 sm:flex sm:flex-wrap items-center gap-2 print:hidden">
                    <select
                      value={filterRegu}
                      onChange={(e) => {
                        setFilterRegu(e.target.value);
                        setFilterPetugas("SEMUA");
                      }}
                      className="text-xs font-semibold bg-white border border-slate-200 rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-[#005C8A]"
                    >
                      <option value="SEMUA">Semua Regu</option>
                      {activeData.daftarRegu.map((rg) => (
                        <option key={rg} value={rg}>Regu {rg}</option>
                      ))}
                    </select>

                    <select
                      value={filterPetugas}
                      onChange={(e) => setFilterPetugas(e.target.value)}
                      className="text-xs font-semibold bg-white border border-slate-200 rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-[#005C8A]"
                    >
                      <option value="SEMUA">Semua Petugas</option>
                      {opsiPetugas.map((ptg: any) => {
                        const val = ptg.kodeAnggota || ptg.idPetugas || "";
                        return (
                          <option key={val} value={val}>
                            {val} (Regu {ptg.regu})
                          </option>
                        );
                      })}
                    </select>

                    <select
                      value={filterKategori}
                      onChange={(e) => setFilterKategori(e.target.value)}
                      className="text-xs font-semibold bg-white border border-slate-200 rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-[#005C8A]"
                    >
                      <option value="SEMUA">Semua Segmen</option>
                      <option value="KCIC">Khusus KCIC</option>
                      <option value="AMR">AMR</option>
                      <option value="NON-AMR">Non-AMR</option>
                    </select>

                    <div className="relative col-span-2 sm:col-span-1">
                      <Search className="w-3.5 h-3.5 absolute left-2.5 top-2.5 text-slate-400" />
                      <input
                        type="text"
                        placeholder="Cari IDPEL / Nama / RBM..."
                        value={searchQuery}
                        onChange={(e) => setSearchQuery(e.target.value)}
                        className="w-full sm:w-44 pl-8 pr-3 py-1.5 bg-white border border-slate-200 rounded-lg text-xs placeholder-slate-400 focus:outline-none focus:border-[#005C8A] font-medium"
                      />
                    </div>
                  </div>
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full text-xs text-left">
                    <thead className="bg-slate-50/80 text-slate-600 font-bold border-b border-slate-100 select-none print:bg-slate-100">
                      <tr>
                        <th className="py-2.5 px-2.5 w-8 text-center print:hidden">Bayar</th>
                        <th className="py-2.5 px-2.5 w-8 text-center">No</th>
                        <th className="py-2.5 px-3 whitespace-nowrap">ID Pelanggan</th>
                        <th
                          className="py-2.5 px-3 cursor-pointer hover:text-[#005C8A] whitespace-nowrap"
                          onClick={() => handleToggleSort("NAMA")}
                        >
                          <div className="flex items-center gap-1">
                            <span>Nama Pelanggan</span>
                            <ArrowUpDown className="w-3 h-3 text-slate-400" />
                          </div>
                        </th>
                        <th className="py-2.5 px-2.5 whitespace-nowrap">Tarif / Daya</th>
                        <th className="py-2.5 px-2.5 font-mono whitespace-nowrap">No. RBM</th>
                        <th
                          className="py-2.5 px-2 text-center cursor-pointer hover:text-[#005C8A]"
                          onClick={() => handleToggleSort("REGU")}
                        >
                          <div className="flex items-center justify-center gap-1">
                            <span>Regu</span>
                            <ArrowUpDown className="w-3 h-3 text-slate-400" />
                          </div>
                        </th>
                        <th
                          className="py-2.5 px-2 text-center cursor-pointer hover:text-[#005C8A]"
                          onClick={() => handleToggleSort("PETUGAS")}
                        >
                          <div className="flex items-center justify-center gap-1">
                            <span>Petugas</span>
                            <ArrowUpDown className="w-3 h-3 text-slate-400" />
                          </div>
                        </th>
                        <th className="py-2.5 px-2 text-center">Segmen</th>
                        <th
                          className="py-2.5 px-3 text-right cursor-pointer hover:text-[#005C8A] whitespace-nowrap"
                          onClick={() => handleToggleSort("NOMINAL")}
                        >
                          <div className="flex items-center justify-end gap-1">
                            <span>Tunggakan (Rp)</span>
                            <ArrowUpDown className="w-3 h-3 text-slate-400" />
                          </div>
                        </th>
                        {activeTab === "LEMBAR_KERJA" && (
                          <>
                            <th className="py-2.5 px-3 text-center w-24 whitespace-nowrap">Status</th>
                            <th className="py-2.5 px-3 text-center w-24 whitespace-nowrap">Paraf</th>
                          </>
                        )}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 text-slate-700">
                      {filteredPelanggan.length > 0 ? (
                        filteredPelanggan.slice(0, activeTab === "RINGKASAN" ? 25 : 120).map((item, idx) => {
                          const isLunas = !!simulasiLunas[item.idpel];
                          return (
                            <tr
                              key={item.idpel}
                              className={`transition ${
                                isLunas ? "bg-emerald-50/50 line-through text-slate-400" : "hover:bg-slate-50/70"
                              }`}
                            >
                              <td className="py-2 px-2.5 text-center print:hidden">
                                <button
                                  onClick={() => {
                                    setSimulasiLunas((prev) => ({ ...prev, [item.idpel]: !prev[item.idpel] }));
                                  }}
                                  className="text-slate-400 hover:text-emerald-600 transition"
                                >
                                  {isLunas ? <CheckSquare className="w-4 h-4 text-emerald-600" /> : <Square className="w-4 h-4" />}
                                </button>
                              </td>
                              <td className="py-2 px-2.5 text-center font-semibold text-slate-400">{idx + 1}</td>
                              <td className="py-2 px-3 font-mono font-bold text-slate-900 whitespace-nowrap">{item.idpel}</td>
                              <td className="py-2 px-3 font-semibold text-slate-800 whitespace-nowrap">{item.nama}</td>
                              <td className="py-2 px-2.5 text-slate-600 whitespace-nowrap">{item.tarifDaya}</td>
                              <td className="py-2 px-2.5 font-mono text-slate-600 whitespace-nowrap">{item.rbm}</td>
                              <td className="py-2 px-2 text-center font-bold text-[#005C8A]">
                                <span className="px-1.5 py-0.5 bg-sky-50 rounded border border-sky-100 font-mono">
                                  {item.regu}
                                </span>
                              </td>
                              <td className="py-2 px-2 text-center font-mono font-semibold text-slate-700">
                                {(item as any).kodeAnggota || item.idPetugas}
                              </td>
                              <td className="py-2 px-2 text-center">
                                {item.kategori === "KCIC" && (
                                  <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-rose-50 text-rose-700 border border-rose-200">
                                    KCIC
                                  </span>
                                )}
                                {item.kategori === "AMR" && (
                                  <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-sky-50 text-[#005C8A] border border-sky-200">
                                    AMR
                                  </span>
                                )}
                                {item.kategori === "NON-AMR" && (
                                  <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-slate-100 text-slate-600 border border-slate-200">
                                    NON-AMR
                                  </span>
                                )}
                              </td>
                              <td className="py-2 px-3 text-right font-black text-slate-900 whitespace-nowrap">
                                Rp {(item.rpTunggakan || 0).toLocaleString("id-ID")}
                              </td>
                              {activeTab === "LEMBAR_KERJA" && (
                                <>
                                  <td className="py-2 px-3 text-center border-l border-slate-200 text-[10px] text-slate-500 whitespace-nowrap">
                                    [ ] Lunas [ ] Segel
                                  </td>
                                  <td className="py-2 px-3 text-center border-l border-slate-200 text-slate-300">
                                    ........
                                  </td>
                                </>
                              )}
                            </tr>
                          );
                        })
                      ) : (
                        <tr>
                          <td colSpan={activeTab === "LEMBAR_KERJA" ? 12 : 10} className="py-8 text-center text-slate-400">
                            Tidak ditemukan data pelanggan sesuai kriteria pencarian.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>

                {/* LEMBAR PENGESAHAN TANDA TANGAN (PRINT PDF) */}
                <div className="hidden print:grid grid-cols-3 gap-6 p-6 mt-8 text-center text-xs text-slate-900 border-t-2 border-slate-800">
                  <div>
                    <p className="font-semibold text-slate-600">Disusun oleh,</p>
                    <p className="font-bold text-slate-900 mt-1">Team Leader Pengelolaan Piutang</p>
                    <div className="h-20" />
                    <p className="font-bold underline text-slate-900">( ............................................ )</p>
                    <p className="text-[10px] text-slate-500 mt-0.5">NIP: .......................................</p>
                  </div>
                  <div>
                    <p className="font-semibold text-slate-600">Diperiksa oleh,</p>
                    <p className="font-bold text-slate-900 mt-1">Asisten Manajer Transaksi Energi</p>
                    <div className="h-20" />
                    <p className="font-bold underline text-slate-900">( ............................................ )</p>
                    <p className="text-[10px] text-slate-500 mt-0.5">NIP: .......................................</p>
                  </div>
                  <div>
                    <p className="font-semibold text-slate-600">Disetujui oleh,</p>
                    <p className="font-bold text-slate-900 mt-1">Manager UP3 Jatinegara</p>
                    <div className="h-20" />
                    <p className="font-bold underline text-slate-900">( ............................................ )</p>
                    <p className="text-[10px] text-slate-500 mt-0.5">NIP: .......................................</p>
                  </div>
                </div>
              </div>
            </>
          )
        )}
      </div>
    </main>
  );
}