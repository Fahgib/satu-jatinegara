import * as XLSX from "xlsx";

export interface PelangganTunggakan {
  idpel: string;
  nama: string;
  tarifDaya: string;
  kategori: "AMR" | "NON-AMR" | "KCIC";
  kelompokTarif: "R" | "B" | "I" | "P" | "T" | "LAINNYA";
  rpTunggakan: number;
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

  let startIndex = 0;
  for (let i = 0; i < Math.min(rows.length, 15); i++) {
    if (rows[i] && rows[i].length > 5) {
      startIndex = i;
      break;
    }
  }

  let nonAmrCount = 0;
  let nonAmrRp = 0;
  let amrBiasaCount = 0;
  let amrBiasaRp = 0;
  let amrKcicCount = 0;
  let amrKcicRp = 0;

  const semuaPelanggan: PelangganTunggakan[] = [];
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
    if (!row || row.length < 5) continue;

    const idpel = String(row[1] || row[2] || "-").trim();
    const kodeMr = String(row[3] || "").trim().toUpperCase();
    const namaPlgn = String(row[4] || "").trim().toUpperCase();
    const tarifDaya = String(row[5] || row[6] || "TARIF REG").trim().toUpperCase();

    let rpRaw = row[13] !== undefined ? row[13] : row[row.length - 3];
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

    // Deteksi Kelompok Tarif
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
      });
    }
  }

  // Urutkan pelanggan dari tunggakan tertinggi
  semuaPelanggan.sort((a, b) => b.rpTunggakan - a.rpTunggakan);

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
  };
};