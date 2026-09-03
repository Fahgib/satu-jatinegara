import pptxgen from "pptxgenjs";
import { HasilKogol } from "./parserKogol";

export const downloadPresentationPptx = (
  data: HasilKogol,
  mode: "BULANAN" | "TAHUNAN",
  allHistory?: Record<string, HasilKogol>
) => {
  const pres = new pptxgen();

  // Dimensi paten Full HD Widescreen 16:9
  pres.defineLayout({ name: "FULL_HD", width: 13.33, height: 7.5 });
  pres.layout = "FULL_HD";

  // ==============================================================
  // PALET WARNA KORPORAT (PLN & Modern Executive)
  // ==============================================================
  const C_NAVY_DARK = "0B192C";
  const C_PLN_BLUE = "005C8A";
  const C_CYAN = "00A3E0";
  const C_ORANGE = "EA580C";
  const C_YELLOW = "FFD100";
  const C_GREEN = "16A34A";
  const C_WHITE = "FFFFFF";
  const C_CARD_BG = "FFFFFF";
  const C_CARD_BORDER = "CBD5E1";
  const C_HEADER_FILL = "F1F5F9";
  const C_TEXT_MAIN = "0F172A";
  const C_TEXT_MUTED = "64748B";
  const C_TEXT_LIGHT = "E2E8F0";
  const C_BG = "EEF2F6";
  const FONT = "Arial";

  // ==============================================================
  // HELPER
  // ==============================================================
  const rp = (n: number) => `Rp ${Math.round(n || 0).toLocaleString("id-ID")}`;
  const num = (n: number) => (n || 0).toLocaleString("id-ID");

  const targetSaldo = 982000000;
  const totalAmrDengan = data.denganKcic.amrRp || 0;
  const totalDengan = data.denganKcic.totalRp || 1;
  const persenAmr = ((totalAmrDengan / totalDengan) * 100).toFixed(1);
  const capaianPersen = (((targetSaldo - data.denganKcic.totalRp) / targetSaldo) * 100).toFixed(1);
  const capaianNaik = data.denganKcic.totalRp > targetSaldo; // true = melebihi target (buruk)

  // ==============================================================
  // SLIDE TUNGGAL: DASHBOARD EKSEKUTIF SATU HALAMAN
  // ==============================================================
  const slide = pres.addSlide();
  slide.background = { color: C_BG };

  // --- aksen garis kiri tipis (identitas korporat) ---
  slide.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: 0.1, h: 7.5, fill: { color: C_YELLOW }, line: { color: C_YELLOW } });
  slide.addShape(pres.ShapeType.rect, { x: 0.1, y: 0, w: 0.1, h: 7.5, fill: { color: C_CYAN }, line: { color: C_CYAN } });

  // ------------------------------------------------------------
  // 1) HEADER (y: 0.3 - 1.15)
  // ------------------------------------------------------------
  slide.addShape(pres.ShapeType.roundRect, {
    x: 0.8, y: 0.3, w: 11.73, h: 0.85,
    fill: { color: C_NAVY_DARK }, line: { color: C_NAVY_DARK },
  });

  slide.addText("⚡ SATU JATINEGARA", {
    x: 1.0, y: 0.35, w: 6.8, h: 0.38,
    fontSize: 18, bold: true, color: C_WHITE, fontFace: FONT,
  });
  slide.addText(
    `LAPORAN KINERJA ${mode}  |  CUT-OFF: ${data.periode.toUpperCase()}  |  PT PLN (PERSERO) UP3 JATINEGARA — UID JAKARTA RAYA`,
    { x: 1.0, y: 0.74, w: 7.3, h: 0.3, fontSize: 8.5, color: C_TEXT_LIGHT, fontFace: FONT, bold: true }
  );

  // Kotak capaian KPI di kanan header
  slide.addShape(pres.ShapeType.roundRect, {
    x: 8.55, y: 0.42, w: 3.83, h: 0.62,
    fill: { color: capaianNaik ? "3F1212" : "0F2E1C" }, line: { color: C_CYAN, width: 0.75 },
  });
  slide.addText(
    [
      { text: "CAPAIAN TARGET:  ", options: { fontSize: 9.5, bold: true, color: C_TEXT_LIGHT } },
      { text: `${capaianPersen}%\n`, options: { fontSize: 13, bold: true, color: capaianNaik ? "F87171" : "4ADE80" } },
      { text: `Target: Rp ${(targetSaldo / 1e6).toFixed(0)} Jt   |   Realisasi: Rp ${(data.denganKcic.totalRp / 1e9).toFixed(2)} M`, options: { fontSize: 7.5, color: "BAE6FD" } },
    ],
    { x: 8.7, y: 0.46, w: 3.55, h: 0.55, fontFace: FONT, valign: "top" }
  );

  // ------------------------------------------------------------
  // 2) KARTU RINGKASAN: DENGAN KCIC / TANPA KCIC (y: 1.25 - 2.55)
  // ------------------------------------------------------------
  const buildSummaryCard = (
    x: number,
    label: string,
    accent: string,
    accentBg: string,
    seg: HasilKogol["denganKcic"]
  ) => {
    slide.addShape(pres.ShapeType.roundRect, { x, y: 1.25, w: 5.7, h: 1.3, fill: { color: C_CARD_BG }, line: { color: accent, width: 1.5 } });
    slide.addShape(pres.ShapeType.rect, { x, y: 1.25, w: 5.7, h: 0.28, fill: { color: accentBg }, line: { color: accentBg } });
    slide.addText(label, { x: x + 0.15, y: 1.28, w: 5.4, h: 0.24, fontSize: 9.5, bold: true, color: accent, fontFace: FONT });

    slide.addText(
      [
        { text: "TOTAL PELANGGAN\n", options: { fontSize: 7.5, color: C_TEXT_MUTED, bold: true } },
        { text: `${num(seg.totalPlgn)} `, options: { fontSize: 19, bold: true, color: C_TEXT_MAIN } },
        { text: "Plgn\n", options: { fontSize: 8.5, color: C_TEXT_MUTED } },
        { text: `AMR: ${num(seg.amrPlgn)}  |  Non-AMR: ${num(seg.nonAmrPlgn)}`, options: { fontSize: 8, color: C_TEXT_MUTED } },
      ],
      { x: x + 0.15, y: 1.58, w: 2.55, h: 0.95, fontFace: FONT, valign: "top" }
    );

    slide.addText(
      [
        { text: "RUPIAH SALDO AKHIR\n", options: { fontSize: 7.5, color: C_TEXT_MUTED, bold: true } },
        { text: `${rp(seg.totalRp)}\n`, options: { fontSize: 13, bold: true, color: accent } },
        { text: `AMR : ${rp(seg.amrRp)}\nNon : ${rp(seg.nonAmrRp)}`, options: { fontSize: 7.5, color: C_TEXT_MUTED } },
      ],
      { x: x + 2.85, y: 1.58, w: 2.7, h: 0.95, fontFace: FONT, valign: "top" }
    );
  };

  buildSummaryCard(0.8, "PERSPEKTIF 1: TOTAL REALISASI (DENGAN KCIC)", C_PLN_BLUE, "E0F2FE", data.denganKcic);
  buildSummaryCard(6.83, "PERSPEKTIF 2: BEBAN REGULER (TANPA KCIC)", C_ORANGE, "FFEDD5", data.tanpaKcic);

  // ------------------------------------------------------------
  // 3) BARIS CHART: 2 DONUT + 1 TREN/PERBANDINGAN (y: 2.65 - 4.65)
  // ------------------------------------------------------------
  const chartCardY = 2.65;
  const chartCardH = 2.0;

  // Donut 1: proporsi pelanggan
  slide.addShape(pres.ShapeType.roundRect, { x: 0.8, y: chartCardY, w: 3.6, h: chartCardH, fill: { color: C_CARD_BG }, line: { color: C_CARD_BORDER, width: 1 } });
  slide.addChart(
    pres.ChartType.doughnut,
    [{ name: "Pelanggan", labels: ["AMR", "NON-AMR"], values: [data.denganKcic.amrPlgn, data.denganKcic.nonAmrPlgn] }],
    {
      x: 0.9, y: 2.78, w: 3.4, h: 1.75,
      chartColors: [C_PLN_BLUE, C_YELLOW],
      showPercent: true, showLegend: true, legendPos: "b", legendFontSize: 7,
      dataLabelFontSize: 8,
      title: "Proporsi Pelanggan (AMR vs Non-AMR)", titleFontSize: 8.5,
    }
  );

  // Donut 2: proporsi rupiah
  slide.addShape(pres.ShapeType.roundRect, { x: 4.55, y: chartCardY, w: 3.6, h: chartCardH, fill: { color: C_CARD_BG }, line: { color: C_CARD_BORDER, width: 1 } });
  slide.addChart(
    pres.ChartType.doughnut,
    [{ name: "Nominal", labels: ["AMR", "NON-AMR"], values: [data.denganKcic.amrRp, data.denganKcic.nonAmrRp] }],
    {
      x: 4.65, y: 2.78, w: 3.4, h: 1.75,
      chartColors: [C_PLN_BLUE, C_ORANGE],
      showPercent: true, showLegend: true, legendPos: "b", legendFontSize: 7,
      dataLabelFontSize: 8,
      title: "Proporsi Nominal (AMR vs Non-AMR)", titleFontSize: 8.5,
    }
  );

  // Panel 3: tren multi-periode (jika ada histori) atau perbandingan periode berjalan
  slide.addShape(pres.ShapeType.roundRect, { x: 8.3, y: chartCardY, w: 4.23, h: chartCardH, fill: { color: C_CARD_BG }, line: { color: C_CARD_BORDER, width: 1 } });

  const hasHistory = allHistory && Object.keys(allHistory).length > 1;
  if (hasHistory) {
    const sortedKeys = Object.keys(allHistory!).sort();
    const chartLabels = sortedKeys.map((k) => allHistory![k].periode);
    const dataDengan = sortedKeys.map((k) => allHistory![k].denganKcic.totalRp);
    const dataTanpa = sortedKeys.map((k) => allHistory![k].tanpaKcic.totalRp);

    slide.addChart(
      pres.ChartType.bar,
      [
        { name: "Dengan KCIC", labels: chartLabels, values: dataDengan },
        { name: "Tanpa KCIC", labels: chartLabels, values: dataTanpa },
      ],
      {
        x: 8.42, y: 2.78, w: 4.0, h: 1.75,
        chartColors: [C_ORANGE, C_PLN_BLUE],
        showLegend: true, legendPos: "t", legendFontSize: 7,
        catAxisLabelFontSize: 7, valAxisLabelFontSize: 7,
        title: "Tren Saldo Multi-Periode", titleFontSize: 8.5,
      }
    );
  } else {
    slide.addChart(
      pres.ChartType.bar,
      [{ name: "Saldo Akhir", labels: ["Dengan KCIC", "Tanpa KCIC"], values: [data.denganKcic.totalRp, data.tanpaKcic.totalRp] }],
      {
        x: 8.42, y: 2.78, w: 4.0, h: 1.75,
        chartColors: [C_PLN_BLUE],
        showLegend: false,
        showValue: true, dataLabelFontSize: 8,
        catAxisLabelFontSize: 7.5, valAxisLabelFontSize: 7,
        title: "Perbandingan Dengan vs Tanpa KCIC", titleFontSize: 8.5,
      }
    );
  }

  // ------------------------------------------------------------
  // 4) TABEL TOP PELANGGAN + RINGKASAN EKSEKUTIF (y: 4.75 - 6.95)
  // ------------------------------------------------------------
  // --- Tabel (kiri) ---
  slide.addShape(pres.ShapeType.roundRect, { x: 0.8, y: 4.75, w: 7.3, h: 2.2, fill: { color: C_CARD_BG }, line: { color: C_CARD_BORDER, width: 1 } });
  slide.addText("PRIORITAS PENAGIHAN: TOP 5 PELANGGAN PARETO", {
    x: 0.95, y: 4.85, w: 7.0, h: 0.25, fontSize: 9.5, bold: true, color: C_PLN_BLUE, fontFace: FONT,
  });

  if (data.semuaPelanggan && data.semuaPelanggan.length > 0) {
    const rows = data.semuaPelanggan.slice(0, 5);
    const tableRows: any[] = [
      [
        { text: "No", options: { bold: true, fill: C_HEADER_FILL, align: "center" } },
        { text: "ID Pelanggan", options: { bold: true, fill: C_HEADER_FILL } },
        { text: "Nama Pelanggan", options: { bold: true, fill: C_HEADER_FILL } },
        { text: "Segmen", options: { bold: true, fill: C_HEADER_FILL, align: "center" } },
        { text: "Nominal (Rp)", options: { bold: true, fill: C_HEADER_FILL, align: "right" } },
        { text: "Kontribusi", options: { bold: true, fill: C_HEADER_FILL, align: "right" } },
      ],
    ];
    rows.forEach((p, idx) => {
      const kontribusi = ((p.rpTunggakan / data.denganKcic.totalRp) * 100).toFixed(2);
      tableRows.push([
        { text: `${idx + 1}`, options: { align: "center" } },
        { text: p.idpel },
        { text: p.nama },
        { text: p.kategori, options: { align: "center" } },
        { text: rp(p.rpTunggakan), options: { align: "right" } },
        { text: `${kontribusi}%`, options: { align: "right" } },
      ]);
    });

    slide.addTable(tableRows as any, {
      x: 0.95, y: 5.15, w: 7.0,
      colW: [0.45, 1.35, 2.3, 0.9, 1.35, 0.65],
      fontSize: 7.5,
      border: { pt: 0.5, color: C_CARD_BORDER },
      fill: { color: C_WHITE },
      color: C_TEXT_MAIN,
      autoPage: false,
      rowH: 0.28,
    });
  } else {
    slide.addText("Data rincian pelanggan tidak tersedia untuk periode ini.", {
      x: 0.95, y: 5.3, w: 7.0, h: 0.4, fontSize: 9, italic: true, color: C_TEXT_MUTED, fontFace: FONT,
    });
  }

  // --- Ringkasan eksekutif (kanan) ---
  slide.addShape(pres.ShapeType.roundRect, { x: 8.25, y: 4.75, w: 4.28, h: 2.2, fill: { color: C_CARD_BG }, line: { color: C_CARD_BORDER, width: 1 } });
  slide.addText("RINGKASAN & STRATEGI EKSEKUTIF", {
    x: 8.4, y: 4.85, w: 4.0, h: 0.25, fontSize: 9.5, bold: true, color: C_TEXT_MAIN, fontFace: FONT,
  });

  slide.addText(
    [
      { text: "📌 ", options: { fontSize: 8 } },
      { text: `Realisasi saldo Rp ${(data.denganKcic.totalRp / 1e9).toFixed(2)} M vs target Rp ${(targetSaldo / 1e6).toFixed(0)} Jt (deviasi ${capaianPersen}%).\n\n`, options: { fontSize: 8, color: C_TEXT_MAIN } },
      { text: "⚠️ ", options: { fontSize: 8 } },
      { text: `${persenAmr}% nominal saldo terkonsentrasi pada kategori AMR — perlu penagihan intensif.\n\n`, options: { fontSize: 8, color: C_TEXT_MAIN } },
      { text: "🚆 ", options: { fontSize: 8 } },
      { text: `Piutang KCIC: ${rp(data.kcicOnly.rp)} (${num(data.kcicOnly.plgn)} IDPEL). Tanpa KCIC, saldo murni unit: Rp ${(data.tanpaKcic.totalRp / 1e6).toFixed(1)} Jt.`, options: { fontSize: 8, color: C_TEXT_MAIN } },
    ],
    { x: 8.4, y: 5.15, w: 4.0, h: 1.75, lineSpacingMultiple: 1.2, fontFace: FONT, valign: "top" }
  );

  // ------------------------------------------------------------
  // Simpan file
  // ------------------------------------------------------------
  pres.writeFile({
    fileName: `SATU_Jatinegara_${mode}_${data.periode.replace(/\s+/g, "_")}.pptx`,
  });
};