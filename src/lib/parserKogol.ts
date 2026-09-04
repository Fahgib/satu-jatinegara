import * as XLSX from "xlsx";

export interface PelangganTunggakan {
  idpel: string;
  nama: string;
  tarifDaya: string;
  kategori: "AMR" | "NON-AMR" | "KCIC";
  kelompokTarif: "R" | "B" | "I" | "P" | "T" | "LAINNYA";
  rpTunggakan: number;
  rbm: string;
  regu: string;
  idPetugas: string;
  kodeAnggota?: string;
  kodePerorangan?: string;
}

export interface RekapMetrik {
  totalPlgn: number;
  totalRp: number;
  amrPlgn: number;
  amrRp: number;
  nonAmrPlgn: number;
  nonAmrRp: number;
}

export interface HasilKogol {
  periode: string;
  kodeBulan: string;
  tahun: string;
  denganKcic: RekapMetrik;
  tanpaKcic: RekapMetrik;
  kcicOnly: { plgn: number; rp: number };
  semuaPelanggan: PelangganTunggakan[];
  rekapTarif: Record<string, { plgn: number; rp: number }>;
  daftarRegu: string[];
  daftarPetugas: { idPetugas: string; regu: string; kodeAnggota?: string }[];
  peringkatRegu: { regu: string; totalRp: number; totalPlgn: number }[];
}

const NAMA_BULAN: Record<string, string> = {
  "01": "Januari",
  "02": "Februari",
  "03": "Maret",
  "04": "April",
  "05": "Mei",
  "06": "Juni",
  "07": "Juli",
  "08": "Agustus",
  "09": "September",
  "10": "Oktober",
  "11": "November",
  "12": "Desember",
};

// Daftar IDPEL resmi KCIC (Di luar daftar ini = NON-KCIC)
const DAFTAR_IDPEL_KCIC = new Set<string>([
  "547104923302",
  "547104923298",
  "544104351417",
  "547104687734",
  "547104960177",
]);

// Daftar kode regu yang valid (2 karakter dari digit ke-4 & 5 RBM)
const DAFTAR_REGU_VALID = new Set<string>(["AA", "AM", "BA", "CA", "DA", "MR"]);

export const parseExcelKogolBuffer = (
  buffer: ArrayBuffer,
  fileName: string = ""
): HasilKogol => {
  const workbook = XLSX.read(buffer, { type: "array" });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  let tahun = "2026";
  let bulan = "08";

  // Pengamanan: pastikan fileName selalu berupa string
  const cleanName = String(fileName || "").toLowerCase();
  const matchFile = cleanName.match(/(20\d{2})(0[1-9]|1[0-2])/);

  const listBulan = [
    { nama: "jan", kode: "01" },
    { nama: "feb", kode: "02" },
    { nama: "mar", kode: "03" },
    { nama: "apr", kode: "04" },
    { nama: "mei", kode: "05" },
    { nama: "may", kode: "05" },
    { nama: "jun", kode: "06" },
    { nama: "jul", kode: "07" },
    { nama: "agu", kode: "08" },
    { nama: "aug", kode: "08" },
    { nama: "sep", kode: "09" },
    { nama: "okt", kode: "10" },
    { nama: "oct", kode: "10" },
    { nama: "nov", kode: "11" },
    { nama: "des", kode: "12" },
    { nama: "dec", kode: "12" },
  ];

  if (matchFile) {
    tahun = matchFile[1];
    bulan = matchFile[2];
  } else if (cleanName) {
    const matchTeks = listBulan.find((b) => cleanName.includes(b.nama));
    if (matchTeks) bulan = matchTeks.kode;
  }

  // Cari Baris Header Kolom
  let headerRowIndex = -1;
  let idxIdpel = -1;
  let idxNama = -1;
  let idxTarif = -1;
  let idxRbm = -1;
  let idxRp = -1;
  let idxMr = -1;

  for (let r = 0; r < Math.min(rows.length, 30); r++) {
    const row = rows[r];
    if (!row || row.length < 3) continue;

    const rowStr = row.map((cell) => String(cell || "").trim().toUpperCase());

    // Deteksi bulan tambahan dari teks header jika dari nama berkas belum terbaca
    if (bulan === "08" && !matchFile) {
      const fullRowText = rowStr.join(" ").toLowerCase();
      const matchHeaderBulan = listBulan.find((b) => fullRowText.includes(b.nama));
      if (matchHeaderBulan) {
        bulan = matchHeaderBulan.kode;
      }
      const matchYear = fullRowText.match(/(20\d{2})/);
      if (matchYear) {
        tahun = matchYear[1];
      }
    }

    const hasIdpel = rowStr.findIndex(
      (h) => h.includes("IDPEL") || h.includes("ID PEL") || h.includes("NO REK") || h.includes("NOREK")
    );
    const hasNama = rowStr.findIndex((h) => h.includes("NAMA"));

    if (hasIdpel !== -1 && hasNama !== -1) {
      headerRowIndex = r;
      idxIdpel = hasIdpel;
      idxNama = hasNama;

      idxRbm = rowStr.findIndex(
        (h) =>
          h.includes("RBM") ||
          h.includes("KODELK") ||
          h.includes("RUTE") ||
          h.includes("KEDUDUKAN") ||
          h.includes("PEMBACA")
      );

      idxTarif = rowStr.findIndex(
        (h) => (h.includes("TARIF") || h.includes("DAYA")) && !h.includes("ALAMAT")
      );
      idxMr = rowStr.findIndex((h) => h === "MR" || h.includes("KODE MR") || h.includes("AMR"));
      idxRp = rowStr.findIndex(
        (h) =>
          h.includes("RP") ||
          h.includes("TAG") ||
          h.includes("SALDO") ||
          h.includes("TOTAL") ||
          h.includes("TUNGGAKAN")
      );
      break;
    }
  }

  const labelPeriode = `${NAMA_BULAN[bulan] || "Agustus"} ${tahun}`;
  const startIndex = headerRowIndex !== -1 ? headerRowIndex + 1 : 1;

  let nonAmrCount = 0;
  let nonAmrRp = 0;
  let amrBiasaCount = 0;
  let amrBiasaRp = 0;
  let amrKcicCount = 0;
  let amrKcicRp = 0;

  const semuaPelanggan: PelangganTunggakan[] = [];
  const setRegu = new Set<string>();
  const mapPetugas = new Map<string, string>();
  const mapPeringkatRegu = new Map<string, { totalRp: number; totalPlgn: number }>();

  const rekapTarif: Record<string, { plgn: number; rp: number }> = {
    R: { plgn: 0, rp: 0 },
    B: { plgn: 0, rp: 0 },
    I: { plgn: 0, rp: 0 },
    P: { plgn: 0, rp: 0 },
    T: { plgn: 0, rp: 0 },
    LAINNYA: { plgn: 0, rp: 0 },
  };

  for (let i = startIndex; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length < 3) continue;

    let idpel = "";
    if (idxIdpel !== -1 && row[idxIdpel]) {
      idpel = String(row[idxIdpel]).trim();
    } else {
      for (let c = 0; c < Math.min(row.length, 4); c++) {
        const clean = String(row[c] || "").replace(/\D/g, "");
        if (clean.length >= 11 && clean.length <= 13) {
          idpel = clean;
          break;
        }
      }
    }

    if (!idpel || idpel.toLowerCase().includes("total") || idpel.toLowerCase().includes("jumlah") || idpel === "-") {
      continue;
    }

    const idpelBersih = idpel.replace(/\D/g, "");
    const kodeMr = String(idxMr !== -1 && row[idxMr] ? row[idxMr] : row[3] || "").trim().toUpperCase();
    const namaPlgn = String(
      (idxNama !== -1 ? row[idxNama] : null) || row[4] || row[2] || ""
    ).trim().toUpperCase();

    // Tarif / Daya
    let tarifDaya = "";
    if (idxTarif !== -1 && row[idxTarif]) {
      tarifDaya = String(row[idxTarif]).trim().toUpperCase();
    }
    if (!tarifDaya || tarifDaya.startsWith("JL") || tarifDaya.startsWith("JALAN") || tarifDaya.length > 20) {
      for (let c = 2; c < Math.min(row.length, 10); c++) {
        const valStr = String(row[c] || "").trim().toUpperCase();
        if (/^(R|B|I|P|T|S)\d/i.test(valStr)) {
          tarifDaya = valStr;
          break;
        }
      }
    }
    if (!tarifDaya) tarifDaya = "TARIF REG";

    // Ekstraksi No RBM
    let rbmRaw = "";
    if (idxRbm !== -1 && row[idxRbm]) {
      rbmRaw = String(row[idxRbm]).trim().toUpperCase();
    }

    if (!rbmRaw || rbmRaw === "-") {
      for (let c = 0; c < row.length; c++) {
        const valStr = String(row[c] || "").trim().toUpperCase();
        if (valStr.startsWith("HAA") && valStr.length >= 7) {
          rbmRaw = valStr;
          break;
        }
      }
    }

    // Ekstraksi Regu & Kode Petugas
    let regu = "-";
    let idPetugas = "-";
    let kodePerorangan = "-";

    if (rbmRaw.length >= 6) {
      const reguCandidate = rbmRaw.substring(3, 5);
      const kodePeroranganCandidate = rbmRaw.charAt(5);
      const kodePetugasCandidate = rbmRaw.substring(3, 6);

      if (DAFTAR_REGU_VALID.has(reguCandidate)) {
        regu = reguCandidate;
        idPetugas = kodePetugasCandidate;
        kodePerorangan = kodePeroranganCandidate;

        setRegu.add(regu);
        mapPetugas.set(idPetugas, regu);
      }
    }

    // Rupiah Tunggakan
    let rp = 0;
    if (idxRp !== -1 && row[idxRp] !== undefined) {
      const rpRaw = row[idxRp];
      rp = typeof rpRaw === "number" ? rpRaw : parseFloat(String(rpRaw).replace(/\./g, "").replace(",", ".")) || 0;
    }
    if (rp === 0) {
      for (let c = row.length - 1; c >= Math.max(0, row.length - 6); c--) {
        const raw = row[c];
        const val = typeof raw === "number" ? raw : parseFloat(String(raw).replace(/\./g, "").replace(",", ".")) || 0;
        if (val > 1000) {
          rp = val;
          break;
        }
      }
    }

    const isAmr = kodeMr === "MR" || String(row[3] || "").toUpperCase() === "MR";
    const isKcic = DAFTAR_IDPEL_KCIC.has(idpelBersih);

    let kategori: "AMR" | "NON-AMR" | "KCIC" = "NON-AMR";
    if (!isAmr) {
      nonAmrCount++;
      nonAmrRp += rp;
      kategori = "NON-AMR";
    } else if (isKcic) {
      amrKcicCount++;
      amrKcicRp += rp;
      kategori = "KCIC";
    } else {
      amrBiasaCount++;
      amrBiasaRp += rp;
      kategori = "AMR";
    }

    let kelompok: "R" | "B" | "I" | "P" | "T" | "LAINNYA" = "LAINNYA";
    const hurufAwalTarif = tarifDaya.charAt(0);
    if (isKcic || hurufAwalTarif === "T") {
      kelompok = "T";
    } else if (["R", "B", "I", "P"].includes(hurufAwalTarif)) {
      kelompok = hurufAwalTarif as any;
    }

    if (rp > 0) {
      rekapTarif[kelompok].plgn += 1;
      rekapTarif[kelompok].rp += rp;

      semuaPelanggan.push({
        idpel,
        nama: namaPlgn,
        tarifDaya,
        kategori,
        kelompokTarif: kelompok,
        rpTunggakan: rp,
        rbm: rbmRaw || "-",
        regu,
        idPetugas,
        kodeAnggota: idPetugas,
        kodePerorangan,
      });

      if (regu !== "-") {
        const existing = mapPeringkatRegu.get(regu) || { totalRp: 0, totalPlgn: 0 };
        existing.totalRp += rp;
        existing.totalPlgn += 1;
        mapPeringkatRegu.set(regu, existing);
      }
    }
  }

  semuaPelanggan.sort((a, b) => b.rpTunggakan - a.rpTunggakan);

  const daftarPetugas = Array.from(mapPetugas.entries())
    .map(([id, rg]) => ({ idPetugas: id, regu: rg, kodeAnggota: id }))
    .sort((a, b) => a.idPetugas.localeCompare(b.idPetugas));

  const peringkatRegu = Array.from(mapPeringkatRegu.entries())
    .map(([rg, data]) => ({ regu: rg, totalRp: data.totalRp, totalPlgn: data.totalPlgn }))
    .sort((a, b) => b.totalRp - a.totalRp);

  return {
    periode: labelPeriode,
    kodeBulan: bulan,
    tahun: tahun,
    denganKcic: {
      totalPlgn: nonAmrCount + amrBiasaCount + amrKcicCount,
      totalRp: nonAmrRp + amrBiasaRp + amrKcicRp,
      amrPlgn: amrBiasaCount + amrKcicCount,
      amrRp: amrBiasaRp + amrKcicRp,
      nonAmrPlgn: nonAmrCount,
      nonAmrRp: nonAmrRp,
    },
    tanpaKcic: {
      totalPlgn: nonAmrCount + amrBiasaCount,
      totalRp: nonAmrRp + amrBiasaRp,
      amrPlgn: amrBiasaCount,
      amrRp: amrBiasaRp,
      nonAmrPlgn: nonAmrCount,
      nonAmrRp: nonAmrRp,
    },
    kcicOnly: {
      plgn: amrKcicCount,
      rp: amrKcicRp,
    },
    semuaPelanggan,
    rekapTarif,
    daftarRegu: Array.from(setRegu).sort(),
    daftarPetugas,
    peringkatRegu,
  };
};