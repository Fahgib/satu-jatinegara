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
  daftarPetugas: { idPetugas: string; regu: string }[];
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

export const parseExcelKogolBuffer = (
  buffer: ArrayBuffer,
  fileName: string = ""
): HasilKogol => {
  const workbook = XLSX.read(buffer, { type: "array" });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  let tahun = "2026";
  let bulan = "06";

  const cleanName = fileName.toLowerCase();
  const matchFile = cleanName.match(/(20\d{2})(0[1-9]|1[0-2])/);

  if (matchFile) {
    tahun = matchFile[1];
    bulan = matchFile[2];
  } else {
    const listBulan = [
      { nama: "jan", kode: "01" },
      { nama: "feb", kode: "02" },
      { nama: "mar", kode: "03" },
      { nama: "apr", kode: "04" },
      { nama: "mei", kode: "05" },
      { nama: "jun", kode: "06" },
      { nama: "jul", kode: "07" },
      { nama: "agu", kode: "08" },
      { nama: "sep", kode: "09" },
      { nama: "okt", kode: "10" },
      { nama: "nov", kode: "11" },
      { nama: "des", kode: "12" },
    ];
    const matchTeks = listBulan.find((b) => cleanName.includes(b.nama));
    if (matchTeks) {
      bulan = matchTeks.kode;
    } else {
      for (let r = 0; r < Math.min(rows.length, 15); r++) {
        const rowStr = JSON.stringify(rows[r] || "");
        const innerMatch = rowStr.match(/(20\d{2})(0[1-9]|1[0-2])/);
        if (innerMatch) {
          tahun = innerMatch[1];
          bulan = innerMatch[2];
          break;
        }
      }
    }
  }

  const labelPeriode = `${NAMA_BULAN[bulan] || "Bulan"} ${tahun}`;

  // 1. Deteksi Baris Header & Letak Kolom RBM Secara Cerdas
  let headerIndex = -1;
  let colIndexRbm = -1;
  let colIndexRp = -1;
  let colIndexIdpel = -1;
  let colIndexNama = -1;
  let colIndexTarif = -1;

  for (let i = 0; i < Math.min(rows.length, 20); i++) {
    const r = rows[i];
    if (!r) continue;
    const rStr = r.map((c) => String(c || "").toUpperCase());

    const rbmIdx = rStr.findIndex((h) =>
      h.includes("RBM") || h.includes("KODELK") || h.includes("PEMBACA") || h.includes("KEDUDUKAN")
    );
    if (rbmIdx !== -1) {
      headerIndex = i;
      colIndexRbm = rbmIdx;
      colIndexIdpel = rStr.findIndex((h) => h.includes("IDPEL") || h.includes("ID PEL"));
      colIndexNama = rStr.findIndex((h) => h.includes("NAMA"));
      colIndexTarif = rStr.findIndex((h) => h.includes("TARIF") || h.includes("DAYA"));
      colIndexRp = rStr.findIndex((h) => h.includes("RP") || h.includes("TAG") || h.includes("SALDO"));
      break;
    }
  }

  const startIndex = headerIndex !== -1 ? headerIndex + 1 : 1;

  let nonAmrCount = 0;
  let nonAmrRp = 0;
  let amrBiasaCount = 0;
  let amrBiasaRp = 0;
  let amrKcicCount = 0;
  let amrKcicRp = 0;

  const semuaPelanggan: PelangganTunggakan[] = [];
  const setRegu = new Set<string>();
  const mapPetugas = new Map<string, string>();

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
    if (!row || row.length < 4) continue;

    const idpel = String(
      (colIndexIdpel !== -1 ? row[colIndexIdpel] : null) || row[1] || row[2] || "-"
    ).trim();

    // Skip baris subtotal / footer
    if (idpel.toLowerCase().includes("total") || idpel.toLowerCase().includes("jumlah")) {
      continue;
    }

    const kodeMr = String(row[3] || "").trim().toUpperCase();
    const namaPlgn = String(
      (colIndexNama !== -1 ? row[colIndexNama] : null) || row[4] || ""
    ).trim().toUpperCase();

    const tarifDaya = String(
      (colIndexTarif !== -1 ? row[colIndexTarif] : null) || row[5] || row[6] || "TARIF REG"
    ).trim().toUpperCase();

    // Pencarian Nilai RBM
    let rbmRaw = "";
    if (colIndexRbm !== -1 && row[colIndexRbm]) {
      rbmRaw = String(row[colIndexRbm]).trim();
    } else {
      // Fallback: telusuri kolom baris yang berisi deretan angka format RBM (panjang 6-12 digit)
      for (let c = 5; c < Math.min(row.length, 14); c++) {
        const valClean = String(row[c] || "").replace(/\D/g, "");
        if (valClean.length >= 6 && valClean.length <= 12) {
          rbmRaw = String(row[c]).trim();
          break;
        }
      }
    }

    // Ekstraksi Regu & ID Petugas
    // Hilangkan semua karakter selain angka agar pemotongan digit presisi
    const rbmDigitsOnly = rbmRaw.replace(/\D/g, "");
    let regu = "-";
    let idPetugas = "-";

    if (rbmDigitsOnly.length >= 6) {
      // Digit ke-4 & ke-5 (indeks 3 s/d 5) -> Regu
      regu = rbmDigitsOnly.substring(3, 5);
      // Digit ke-4, ke-5 & ke-6 (indeks 3 s/d 6) -> ID Petugas
      idPetugas = rbmDigitsOnly.substring(3, 6);

      setRegu.add(regu);
      mapPetugas.set(idPetugas, regu);
    }

    // Rupiah Tunggakan
    let rpRaw =
      colIndexRp !== -1 && row[colIndexRp] !== undefined
        ? row[colIndexRp]
        : row[13] !== undefined
        ? row[13]
        : row[row.length - 3];

    let rp =
      typeof rpRaw === "number"
        ? rpRaw
        : parseFloat(String(rpRaw).replace(/\./g, "").replace(",", ".")) || 0;

    const isAmr = kodeMr === "MR";
    const isKcic = namaPlgn.includes("KCIC");

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
      });
    }
  }

  semuaPelanggan.sort((a, b) => b.rpTunggakan - a.rpTunggakan);

  const daftarPetugas = Array.from(mapPetugas.entries())
    .map(([id, rg]) => ({ idPetugas: id, regu: rg }))
    .sort((a, b) => a.idPetugas.localeCompare(b.idPetugas));

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
  };
};